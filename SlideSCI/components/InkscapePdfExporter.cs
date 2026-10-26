using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Win32;

namespace SlideSCI
{
    /// <summary>
    /// PowerPoint 导出的 SVG 版式与透明度正确，但内嵌图片会被降采样并转为 JPEG。
    /// 这里把与原图内容一致的内嵌图片换回原始数据，再交给 Inkscape 输出 PDF。
    /// </summary>
    public static class InkscapePdfExporter
    {
        private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
        private static readonly XNamespace XLink = "http://www.w3.org/1999/xlink";
        private const int FingerprintSize = 16;
        // 灰度指纹平均差异（0-255）。带亮度、重新着色等效果的图片差异明显，保留 PowerPoint 渲染结果。
        private const double MaxFingerprintDifference = 10;
        private const double MaxAspectDifference = 0.02;
        // 高分辨率 PNG 与 SVG 内嵌图的光栅化方式不同，边缘会差 1~2 像素，比例与差异阈值都放宽一些。
        private const double MaxRenderAspectDifference = 0.04;
        private const double MaxRenderDifference = 12;
        private const byte OpaqueAlpha = 8;

        public static string FindInkscape(string configuredPath)
        {
            if (IsInkscape(configuredPath)) return configuredPath;
            foreach (RegistryHive hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            {
                foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    try
                    {
                        using (RegistryKey root = RegistryKey.OpenBaseKey(hive, view))
                        using (RegistryKey key = root.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\inkscape.exe"))
                        {
                            string path = (key?.GetValue(null) as string)?.Trim('"');
                            if (IsInkscape(path)) return path;
                        }
                    }
                    catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
                    {
                        Debug.WriteLine($"读取 Inkscape 注册表路径失败：{ex.Message}");
                    }
                }
            }
            var candidates = (Environment.GetEnvironmentVariable("PATH") ?? "")
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(directory => directory.Trim().Trim('"'))
                .Concat(new[]
                {
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    Environment.GetEnvironmentVariable("ProgramW6432"),
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs"),
                }.Where(directory => !string.IsNullOrEmpty(directory))
                 .Select(directory => Path.Combine(directory, "Inkscape", "bin")));
            foreach (string directory in candidates)
            {
                try
                {
                    string path = Path.Combine(directory, "inkscape.exe");
                    if (IsInkscape(path)) return path;
                }
                catch (ArgumentException) { }
            }
            return null;
        }

        public static bool IsInkscape(string path) =>
            !string.IsNullOrWhiteSpace(path) && File.Exists(path) &&
            Path.GetFileNameWithoutExtension(path).Equals("inkscape", StringComparison.OrdinalIgnoreCase);

        public static int CountEmbeddedImages(XDocument svg) => EmbeddedImages(svg).Count();

        public sealed class ReplacementResult
        {
            public int Total { get; internal set; }
            public int Originals { get; internal set; }
            public int Rerendered { get; internal set; }
        }

        /// <summary>
        /// 先把内容与原图一致的内嵌图片换回原图；带三维旋转等效果、无法对应原图的图片，
        /// 再用同一形状的高分辨率 PNG 渲染替换。
        /// </summary>
        public static ReplacementResult ReplaceEmbeddedImages(XDocument svg, IEnumerable<byte[]> originalImages,
            IEnumerable<byte[]> shapeRenders = null)
        {
            var result = new ReplacementResult();
            List<OriginalImage> originals = originalImages
                .Select(OriginalImage.TryCreate)
                .Where(original => original != null)
                .ToList();
            var unmatched = new List<XAttribute>();
            foreach (XElement image in EmbeddedImages(svg).ToList())
            {
                result.Total++;
                XAttribute href = image.Attribute(XLink + "href") ?? image.Attribute("href");
                byte[] renditionData = DecodeDataUri(href.Value);
                if (renditionData == null) continue;
                Fingerprint rendition;
                try
                {
                    using (var stream = new MemoryStream(renditionData))
                    using (var bitmap = new Bitmap(stream))
                    {
                        rendition = Fingerprint.From(bitmap);
                    }
                }
                catch (ArgumentException) { continue; }

                OriginalImage match = originals
                    .Where(original => original.PixelCount > rendition.PixelCount &&
                        Math.Abs(original.Fingerprint.Aspect / rendition.Aspect - 1) <= MaxAspectDifference)
                    .Select(original => new { original, difference = original.Fingerprint.DifferenceFrom(rendition) })
                    .Where(candidate => candidate.difference <= MaxFingerprintDifference)
                    .OrderBy(candidate => candidate.difference)
                    .Select(candidate => candidate.original)
                    .FirstOrDefault();
                if (match == null) { unmatched.Add(href); continue; }
                // width/height 属性与引用它的 transform 不变，原图只是在同一区域内以更高分辨率绘制。
                href.Value = match.DataUri;
                result.Originals++;
            }
            if (unmatched.Count > 0 && shapeRenders != null)
                result.Rerendered = ReplaceWithRenders(unmatched, shapeRenders);
            return result;
        }

        private static int ReplaceWithRenders(List<XAttribute> unmatched, IEnumerable<byte[]> shapeRenders)
        {
            var renders = new List<ArgbImage>();
            try
            {
                foreach (byte[] data in shapeRenders)
                {
                    ArgbImage render = ArgbImage.TryDecode(data);
                    if (render != null && render.Opaque.Width > 0) renders.Add(render);
                }
                int replaced = 0;
                foreach (XAttribute href in unmatched)
                {
                    ArgbImage rendition = ArgbImage.TryDecode(DecodeDataUri(href.Value));
                    if (rendition == null || rendition.Opaque.Width == 0) continue;
                    double renditionAspect = rendition.Opaque.Width / (double)rendition.Opaque.Height;
                    var match = renders
                        .Where(render => (long)render.Opaque.Width * render.Opaque.Height >
                            (long)rendition.Opaque.Width * rendition.Opaque.Height &&
                            Math.Abs(render.Opaque.Width / (double)render.Opaque.Height / renditionAspect - 1) <= MaxRenderAspectDifference)
                        .Select(render => new { render, difference = render.DifferenceFrom(rendition) })
                        .Where(candidate => candidate.difference <= MaxRenderDifference)
                        .OrderBy(candidate => candidate.difference)
                        .FirstOrDefault();
                    if (match == null) continue;
                    href.Value = "data:image/png;base64," + Convert.ToBase64String(match.render.FitInto(rendition));
                    replaced++;
                }
                return replaced;
            }
            finally
            {
                foreach (ArgbImage render in renders) render.Dispose();
            }
        }

        public static void ConvertSvgToPdf(string inkscapePath, string svgPath, string pdfPath, int timeoutMilliseconds = 180000)
        {
            // inkscape.com 是控制台入口，重定向输出更可靠；旧版只有 inkscape.exe。
            string console = Path.ChangeExtension(inkscapePath, ".com");
            var startInfo = new ProcessStartInfo
            {
                FileName = File.Exists(console) ? console : inkscapePath,
                Arguments = $"--export-type=pdf --export-filename=\"{pdfPath}\" \"{svgPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = Path.GetDirectoryName(svgPath),
            };
            var output = new StringBuilder();
            using (var process = new Process { StartInfo = startInfo })
            {
                process.OutputDataReceived += (s, e) => { if (e.Data != null) lock (output) output.AppendLine(e.Data); };
                process.ErrorDataReceived += (s, e) => { if (e.Data != null) lock (output) output.AppendLine(e.Data); };
                if (File.Exists(pdfPath)) File.Delete(pdfPath);
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                if (!process.WaitForExit(timeoutMilliseconds))
                {
                    try { process.Kill(); } catch (InvalidOperationException) { }
                    throw new TimeoutException("Inkscape 转换超时。");
                }
                process.WaitForExit();
                if (process.ExitCode != 0 || !File.Exists(pdfPath))
                {
                    string log;
                    lock (output) log = output.ToString().Trim();
                    if (log.Length > 1500) log = log.Substring(log.Length - 1500);
                    throw new InvalidOperationException($"Inkscape 转换失败（退出码 {process.ExitCode}）。\n{log}");
                }
            }
        }

        private static IEnumerable<XElement> EmbeddedImages(XDocument svg) =>
            svg.Descendants(Svg + "image").Where(image =>
                ((string)image.Attribute(XLink + "href") ?? (string)image.Attribute("href") ?? "")
                    .StartsWith("data:image/", StringComparison.OrdinalIgnoreCase));

        private static byte[] DecodeDataUri(string uri)
        {
            int comma = uri.IndexOf(',');
            if (comma < 0 || uri.LastIndexOf(";base64", comma, StringComparison.OrdinalIgnoreCase) < 0) return null;
            try { return Convert.FromBase64String(uri.Substring(comma + 1)); }
            catch (FormatException) { return null; }
        }

        /// <summary>32 位 ARGB 图片及其不透明区域；PowerPoint 的 PNG 渲染四周带透明留白，按不透明区域对齐。</summary>
        private sealed class ArgbImage : IDisposable
        {
            public Bitmap Bitmap { get; }
            public Rectangle Opaque { get; }
            public int Width => Bitmap.Width;
            public int Height => Bitmap.Height;

            private ArgbImage(Bitmap bitmap)
            {
                Bitmap = bitmap;
                Opaque = FindOpaqueBounds(bitmap);
            }

            public static ArgbImage TryDecode(byte[] data)
            {
                if (data == null || data.Length == 0) return null;
                try
                {
                    using (var stream = new MemoryStream(data))
                    using (var source = new Bitmap(stream))
                    {
                        var bitmap = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
                        using (Graphics graphics = Graphics.FromImage(bitmap))
                        {
                            graphics.Clear(Color.Transparent);
                            graphics.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
                        }
                        return new ArgbImage(bitmap);
                    }
                }
                catch (Exception ex) when (ex is ArgumentException || ex is ExternalException || ex is OutOfMemoryException)
                {
                    Debug.WriteLine($"无法读取渲染图片：{ex.Message}");
                    return null;
                }
            }

            /// <summary>把本图的不透明区域缩放到对方不透明区域的大小，叠加白底后比较灰度平均差异（0-255）。</summary>
            public double DifferenceFrom(ArgbImage other)
            {
                Rectangle target = new Rectangle(Point.Empty, other.Opaque.Size);
                using (Bitmap mine = Flatten(Bitmap, Opaque, target.Size))
                using (Bitmap theirs = Flatten(other.Bitmap, other.Opaque, target.Size))
                {
                    int[] a = ReadPixels(mine), b = ReadPixels(theirs);
                    long total = 0;
                    for (int i = 0; i < a.Length; i++) total += Math.Abs(Gray(a[i]) - Gray(b[i]));
                    return total / (double)a.Length;
                }
            }

            /// <summary>按对方的画布布局输出高分辨率 PNG：不透明内容 1:1 放在对方不透明区域按比例放大后的位置。</summary>
            public byte[] FitInto(ArgbImage rendition)
            {
                double scaleX = Opaque.Width / (double)rendition.Opaque.Width;
                double scaleY = Opaque.Height / (double)rendition.Opaque.Height;
                int canvasWidth = Math.Max(Opaque.Width, (int)Math.Round(rendition.Width * scaleX));
                int canvasHeight = Math.Max(Opaque.Height, (int)Math.Round(rendition.Height * scaleY));
                int left = Math.Min(canvasWidth - Opaque.Width, (int)Math.Round(rendition.Opaque.X * scaleX));
                int top = Math.Min(canvasHeight - Opaque.Height, (int)Math.Round(rendition.Opaque.Y * scaleY));
                using (var canvas = new Bitmap(canvasWidth, canvasHeight, PixelFormat.Format32bppArgb))
                {
                    using (Graphics graphics = Graphics.FromImage(canvas))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.CompositingMode = CompositingMode.SourceCopy;
                        graphics.DrawImage(Bitmap, new Rectangle(left, top, Opaque.Width, Opaque.Height),
                            Opaque, GraphicsUnit.Pixel);
                    }
                    using (var png = new MemoryStream())
                    {
                        canvas.Save(png, ImageFormat.Png);
                        return png.ToArray();
                    }
                }
            }

            public void Dispose() => Bitmap.Dispose();

            private static Bitmap Flatten(Bitmap source, Rectangle region, Size size)
            {
                var result = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
                using (Graphics graphics = Graphics.FromImage(result))
                {
                    graphics.Clear(Color.White);
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.DrawImage(source, new Rectangle(Point.Empty, size), region, GraphicsUnit.Pixel);
                }
                return result;
            }

            private static Rectangle FindOpaqueBounds(Bitmap bitmap)
            {
                int[] pixels = ReadPixels(bitmap);
                int minX = bitmap.Width, minY = bitmap.Height, maxX = -1, maxY = -1;
                for (int y = 0; y < bitmap.Height; y++)
                {
                    int row = y * bitmap.Width;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if ((byte)(pixels[row + x] >> 24) <= OpaqueAlpha) continue;
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
                return maxX < 0 ? Rectangle.Empty : Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            }

            private static int[] ReadPixels(Bitmap bitmap)
            {
                var pixels = new int[bitmap.Width * bitmap.Height];
                BitmapData data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    for (int y = 0; y < bitmap.Height; y++)
                        Marshal.Copy(data.Scan0 + y * data.Stride, pixels, y * bitmap.Width, bitmap.Width);
                }
                finally { bitmap.UnlockBits(data); }
                return pixels;
            }

            private static int Gray(int argb) =>
                (((argb >> 16) & 0xFF) * 299 + ((argb >> 8) & 0xFF) * 587 + (argb & 0xFF) * 114) / 1000;
        }

        private sealed class OriginalImage
        {
            public Fingerprint Fingerprint { get; }
            public long PixelCount => Fingerprint.PixelCount;
            private readonly byte[] data;
            private readonly string mimeType;
            private string dataUri;

            public string DataUri => dataUri ?? (dataUri = "data:" + mimeType + ";base64," + Convert.ToBase64String(data));

            private OriginalImage(byte[] data, string mimeType, Fingerprint fingerprint)
            {
                this.data = data;
                this.mimeType = mimeType;
                Fingerprint = fingerprint;
            }

            public static OriginalImage TryCreate(byte[] data)
            {
                if (data == null || data.Length == 0) return null;
                try
                {
                    using (var stream = new MemoryStream(data))
                    using (Image image = Image.FromStream(stream))
                    {
                        // 矢量原图（EMF/WMF）交给 PowerPoint 的渲染结果处理。
                        if (!(image is Bitmap bitmap)) return null;
                        var fingerprint = Fingerprint.From(bitmap);
                        if (image.RawFormat.Guid == ImageFormat.Jpeg.Guid)
                            return new OriginalImage(data, "image/jpeg", fingerprint);
                        if (image.RawFormat.Guid == ImageFormat.Png.Guid)
                            return new OriginalImage(data, "image/png", fingerprint);
                        // BMP、GIF、TIFF 等无损转为 PNG，保证 Inkscape 可读取。
                        using (var png = new MemoryStream())
                        {
                            bitmap.Save(png, ImageFormat.Png);
                            return new OriginalImage(png.ToArray(), "image/png", fingerprint);
                        }
                    }
                }
                catch (Exception ex) when (ex is ArgumentException || ex is ExternalException || ex is OutOfMemoryException)
                {
                    Debug.WriteLine($"无法读取原始图片：{ex.Message}");
                    return null;
                }
            }
        }

        private sealed class Fingerprint
        {
            public double Aspect { get; private set; }
            public long PixelCount { get; private set; }
            private readonly byte[] gray = new byte[FingerprintSize * FingerprintSize];

            public static Fingerprint From(Bitmap source)
            {
                var result = new Fingerprint
                {
                    Aspect = source.Width / (double)source.Height,
                    PixelCount = (long)source.Width * source.Height,
                };
                // 统一叠加到白底后比较，透明 PNG 与 JPEG 渲染结果可以直接对照。
                using (var thumbnail = new Bitmap(FingerprintSize, FingerprintSize, PixelFormat.Format24bppRgb))
                {
                    using (Graphics graphics = Graphics.FromImage(thumbnail))
                    {
                        graphics.Clear(Color.White);
                        graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        graphics.DrawImage(source, new Rectangle(0, 0, FingerprintSize, FingerprintSize));
                    }
                    for (int y = 0; y < FingerprintSize; y++)
                        for (int x = 0; x < FingerprintSize; x++)
                        {
                            Color color = thumbnail.GetPixel(x, y);
                            result.gray[y * FingerprintSize + x] = (byte)((color.R * 299 + color.G * 587 + color.B * 114) / 1000);
                        }
                }
                return result;
            }

            public double DifferenceFrom(Fingerprint other)
            {
                long total = 0;
                for (int i = 0; i < gray.Length; i++) total += Math.Abs(gray[i] - other.gray[i]);
                return total / (double)gray.Length;
            }
        }
    }
}
