using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        // PpShapeFormat 未公开 SVG 成员，ShapeRange.Export 传入 6 时输出 SVG。
        private const PowerPoint.PpShapeFormat ShapeFormatSvg = (PowerPoint.PpShapeFormat)6;

        private void HighQualityPdf_Click(object sender, RibbonControlEventArgs e)
        {
            string workDirectory = null;
            string step = "检查所选对象";
            try
            {
                if (app.Windows.Count == 0 || app.ActiveWindow == null ||
                    app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                { MessageBox.Show("请先选中要导出的图片或形状。", "提示"); return; }
                var selection = app.ActiveWindow.Selection;
                PowerPoint.ShapeRange shapes = selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange;
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) { MessageBox.Show("请在普通视图中选择幻灯片上的对象。", "提示"); return; }

                step = "查找 Inkscape";
                string inkscape = InkscapePdfExporter.FindInkscape(Properties.Settings.Default.InkscapePath) ?? BrowseForInkscape();
                if (inkscape == null) return;

                string pdfPath = AskHighQualityPdfPath(shapes);
                if (pdfPath == null) return;

                Cursor.Current = Cursors.WaitCursor;
                workDirectory = Path.Combine(Path.GetTempPath(), "SlideSCI-pdf-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(workDirectory);
                string svgPath = Path.Combine(workDirectory, "selection.svg");

                step = "导出 SVG";
                shapes.Export(svgPath, ShapeFormatSvg);
                XDocument svg = XDocument.Load(svgPath);
                int imageCount = InkscapePdfExporter.CountEmbeddedImages(svg);
                InkscapePdfExporter.ReplacementResult replaced = null;
                if (imageCount > 0)
                {
                    step = "替换高分辨率图片";
                    // 渲染是延迟执行的：只有存在无法对应原图的图片时才逐个导出图片形状。
                    replaced = InkscapePdfExporter.ReplaceEmbeddedImages(svg,
                        ReadSlideImages(app.ActivePresentation, slide, workDirectory),
                        RenderPictureShapes(shapes, app.ActivePresentation.PageSetup, workDirectory));
                    svg.Save(svgPath, SaveOptions.DisableFormatting);
                }

                step = "Inkscape 转换 PDF";
                InkscapePdfExporter.ConvertSvgToPdf(inkscape, svgPath, pdfPath);
                Cursor.Current = Cursors.Default;

                string message = $"已导出到：\n{pdfPath}";
                if (replaced != null)
                {
                    message += $"\n\n共 {replaced.Total} 张图片：{replaced.Originals} 张使用原始图片";
                    if (replaced.Rerendered > 0)
                        message += $"，{replaced.Rerendered} 张带三维旋转等效果的图片按 {RenderScale} 倍分辨率重新渲染";
                    message += "。";
                    int kept = replaced.Total - replaced.Originals - replaced.Rerendered;
                    if (kept > 0)
                        message += $"\n其余 {kept} 张已是原始分辨率或无法对应，保留 PowerPoint 的渲染结果。";
                }
                if (MessageBox.Show(message + "\n\n是否打开 PDF？", "导出完成",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    Process.Start(new ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                Debug.WriteLine(ex);
                MessageBox.Show($"导出高清 PDF 时出错（{step}）：{ex.Message}", "操作失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (workDirectory != null)
                {
                    try { Directory.Delete(workDirectory, true); }
                    catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    { Debug.WriteLine($"清理高清 PDF 临时文件时出错: {ex.Message}"); }
                }
            }
        }

        private const int RenderScale = 4;
        private const float MaxRenderPixels = 6000;

        /// <summary>
        /// 逐个导出选中（含组合内）的图片形状为高分辨率 PNG，用于替换 SVG 中带效果、无法对应原图的低分辨率位图。
        /// </summary>
        private static IEnumerable<byte[]> RenderPictureShapes(PowerPoint.ShapeRange shapes,
            PowerPoint.PageSetup pageSetup, string workDirectory)
        {
            var targets = new List<PowerPoint.ShapeRange>();
            for (int i = 1; i <= shapes.Count; i++) CollectPictureShapes(SingleShapeRange(shapes[i]), targets);
            int index = 0;
            foreach (PowerPoint.ShapeRange target in targets)
            {
                byte[] data = null;
                try
                {
                    // ppRelativeToSlide：宽高参数是整张幻灯片的像素尺寸，传入点数为 96 dpi。
                    float longest = Math.Max(target.Width, target.Height) * 96f / 72f;
                    float scale = Math.Min(RenderScale, MaxRenderPixels / Math.Max(1f, longest));
                    if (scale <= 1.2f) continue;
                    string path = Path.Combine(workDirectory, $"render-{index++}.png");
                    target.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                        Math.Max(1, (int)Math.Round(pageSetup.SlideWidth * scale)),
                        Math.Max(1, (int)Math.Round(pageSetup.SlideHeight * scale)),
                        PowerPoint.PpExportMode.ppRelativeToSlide);
                    data = File.ReadAllBytes(path);
                }
                catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException || ex is IOException)
                {
                    Debug.WriteLine($"渲染图片形状失败：{ex.Message}");
                }
                if (data != null) yield return data;
            }
        }

        /// <summary>ShapeRange.Export 只能作用于 ShapeRange；按 Id 找到形状在所属集合中的序号再取单元素范围。</summary>
        private static PowerPoint.ShapeRange SingleShapeRange(PowerPoint.Shape shape)
        {
            if (shape.Child == Office.MsoTriState.msoTrue)
            {
                PowerPoint.GroupShapes items = shape.ParentGroup.GroupItems;
                for (int i = 1; i <= items.Count; i++)
                    if (items[i].Id == shape.Id) return items.Range(i);
            }
            else
            {
                PowerPoint.Shapes items = ((PowerPoint.Slide)shape.Parent).Shapes;
                for (int i = 1; i <= items.Count; i++)
                    if (items[i].Id == shape.Id) return items.Range(i);
            }
            throw new InvalidOperationException($"无法定位形状 {shape.Name}。");
        }

        private static void CollectPictureShapes(PowerPoint.ShapeRange range, List<PowerPoint.ShapeRange> targets)
        {
            PowerPoint.Shape shape = range[1];
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                for (int i = 1; i <= shape.GroupItems.Count; i++)
                    CollectPictureShapes(shape.GroupItems.Range(i), targets);
                return;
            }
            bool isPicture = shape.Type == Office.MsoShapeType.msoPicture ||
                shape.Type == Office.MsoShapeType.msoLinkedPicture;
            if (!isPicture && shape.Type == Office.MsoShapeType.msoPlaceholder)
            {
                try { isPicture = shape.PlaceholderFormat.ContainedType == Office.MsoShapeType.msoPicture; }
                catch (System.Runtime.InteropServices.COMException) { }
            }
            if (!isPicture)
            {
                try { isPicture = shape.Fill.Type == Office.MsoFillType.msoFillPicture; }
                catch (System.Runtime.InteropServices.COMException) { }
            }
            if (isPicture) targets.Add(range);
        }

        private string BrowseForInkscape()
        {
            if (MessageBox.Show("未找到 Inkscape。高清 PDF 需要 Inkscape 1.0 或更高版本。\n\n是否手动选择 inkscape.exe？",
                "未找到 Inkscape", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return null;
            using (var dialog = new OpenFileDialog
            {
                Title = "选择 inkscape.exe",
                Filter = "Inkscape (inkscape.exe;inkscape.com)|inkscape.exe;inkscape.com",
            })
            {
                if (dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) != DialogResult.OK) return null;
                Properties.Settings.Default.InkscapePath = dialog.FileName;
                try { Properties.Settings.Default.Save(); }
                catch (Exception ex) { Debug.WriteLine($"保存 Inkscape 路径失败：{ex.Message}"); }
                return dialog.FileName;
            }
        }

        private string AskHighQualityPdfPath(PowerPoint.ShapeRange shapes)
        {
            var presentation = app.ActivePresentation;
            string name = shapes.Count == 1 ? shapes[1].Name
                : Path.GetFileNameWithoutExtension(presentation.Name) + "_选中对象";
            name = string.Concat(name.Split(Path.GetInvalidFileNameChars())).Trim();
            using (var dialog = new SaveFileDialog
            {
                Title = "导出高清 PDF",
                Filter = "PDF 文件 (*.pdf)|*.pdf",
                FileName = string.IsNullOrEmpty(name) ? "导出" : name,
                OverwritePrompt = true,
            })
            {
                // Path 可能是云端 URL，仅在本地目录存在时使用。
                string directory = presentation.Path;
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) dialog.InitialDirectory = directory;
                return dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>保存本地副本，读取当前幻灯片引用的全部原始图片数据，不修改原文稿及其保存状态。</summary>
        private static List<byte[]> ReadSlideImages(PowerPoint.Presentation presentation, PowerPoint.Slide slide, string workDirectory)
        {
            string copyPath = Path.Combine(workDirectory, "copy.pptx");
            presentation.SaveCopyAs(copyPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            uint slideId = (uint)slide.SlideID;
            var images = new List<byte[]>();
            using (PresentationDocument document = PresentationDocument.Open(copyPath, false))
            {
                PresentationPart presentationPart = document.PresentationPart;
                var slideIdEntry = presentationPart?.Presentation.SlideIdList?.Elements<DocumentFormat.OpenXml.Presentation.SlideId>()
                    .FirstOrDefault(s => s.Id != null && s.Id.Value == slideId);
                if (slideIdEntry?.RelationshipId == null) return images;
                if (!(presentationPart.GetPartById(slideIdEntry.RelationshipId.Value) is SlidePart slidePart)) return images;
                foreach (ImagePart imagePart in slidePart.ImageParts.Distinct())
                {
                    using (Stream stream = imagePart.GetStream())
                    using (var buffer = new MemoryStream())
                    {
                        stream.CopyTo(buffer);
                        images.Add(buffer.ToArray());
                    }
                }
            }
            return images;
        }
    }
}
