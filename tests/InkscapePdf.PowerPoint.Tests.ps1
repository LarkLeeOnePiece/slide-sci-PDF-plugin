# Run with Windows PowerShell -STA. Needs PowerPoint, Inkscape 1.x and the Visual Studio C# compiler.
# Compiles only InkscapePdfExporter.cs, exports shapes from a scratch presentation and checks the PDF images.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repoPath = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repoPath 'artifacts/diagnostics/inkscape-pdf'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$compiler = & $vswhere -latest -products * -find 'MSBuild\**\Roslyn\csc.exe' | Select-Object -First 1
if (-not $compiler) { throw 'The test needs the Visual Studio C# compiler.' }
$framework = 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2'
$assembly = Join-Path $outputDirectory ('InkscapePdfExporter-' + [Guid]::NewGuid().ToString('N') + '.dll')
$references = @('mscorlib.dll','System.dll','System.Core.dll','System.Drawing.dll','System.Xml.dll','System.Xml.Linq.dll') |
    ForEach-Object { '/reference:' + (Join-Path $framework $_) }
& $compiler /nologo /noconfig /nostdlib+ /langversion:7.3 /target:library ('/out:' + $assembly) $references `
    (Join-Path $repoPath 'SlideSCI/components/InkscapePdfExporter.cs')
if ($LASTEXITCODE -ne 0) { throw 'InkscapePdfExporter.cs did not compile.' }
Add-Type -Path $assembly
$inkscape = [SlideSCI.InkscapePdfExporter]::FindInkscape($null)
if (-not $inkscape) { throw 'Inkscape was not found.' }

$photo = Join-Path $outputDirectory 'photo.png'
$bitmap = New-Object Drawing.Bitmap 2000, 1500
$graphics = [Drawing.Graphics]::FromImage($bitmap)
$brush = New-Object Drawing.Drawing2D.LinearGradientBrush((New-Object Drawing.Rectangle 0, 0, 2000, 1500), [Drawing.Color]::Red, [Drawing.Color]::Blue, 45)
$graphics.FillRectangle($brush, 0, 0, 2000, 1500)
for ($x = 0; $x -lt 2000; $x += 4) { $graphics.DrawLine([Drawing.Pens]::White, $x, 1200, $x, 1500) }
$graphics.Dispose(); $bitmap.Save($photo); $bitmap.Dispose()
$alpha = Join-Path $outputDirectory 'alpha.png'
$bitmap = New-Object Drawing.Bitmap 1000, 1000, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics = [Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([Drawing.Color]::Transparent); $graphics.FillEllipse([Drawing.Brushes]::Orange, 100, 100, 800, 800)
$graphics.Dispose(); $bitmap.Save($alpha); $bitmap.Dispose()

$createdApp = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $createdApp = $true }
$presentation = $null
try {
    $presentation = $app.Presentations.Add(0)
    $slide = $presentation.Slides.Add(1, 12)
    $rotated = $slide.Shapes.AddPicture($photo, 0, -1, 40, 40, 300, 225); $rotated.Rotation = 15
    $brightened = $slide.Shapes.AddPicture($photo, 0, -1, 380, 40, 160, 120); $brightened.PictureFormat.Brightness = 0.8
    $transparent = $slide.Shapes.AddPicture($alpha, 0, -1, 560, 40, 100, 100)
    $rectangle = $slide.Shapes.AddShape(1, 380, 200, 120, 80)
    $rectangle.Fill.ForeColor.RGB = 0x00AA00; $rectangle.Fill.Transparency = 0.5
    # Windows PowerShell reads BOM-less scripts as ANSI, so build the non-ASCII caption from code points.
    $caption = 'Caption ' + [char]0x3B1 + ' ' + [char]0x3B2 + ' ' + [char]0x4E2D + [char]0x6587
    $text = $slide.Shapes.AddTextbox(1, 40, 300, 300, 40); $text.TextFrame.TextRange.Text = $caption
    $perspective = $slide.Shapes.AddPicture($photo, 0, -1, 560, 200, 120, 90)
    $perspective.ThreeD.SetPresetCamera(26); $perspective.ThreeD.RotationY = 30

    $svgPath = Join-Path $outputDirectory 'selection.svg'
    $slide.Shapes.Range([object[]]@(1, 2, 3, 4, 5, 6)).Export($svgPath, 6)
    $copyPath = Join-Path $outputDirectory 'copy.pptx'
    $presentation.SaveCopyAs($copyPath, 24, 0)
    $zip = [IO.Compression.ZipFile]::OpenRead($copyPath)
    $images = New-Object 'System.Collections.Generic.List[byte[]]'
    foreach ($entry in $zip.Entries | Where-Object FullName -like 'ppt/media/*') {
        $stream = $entry.Open(); $buffer = New-Object IO.MemoryStream; $stream.CopyTo($buffer); $stream.Dispose()
        $images.Add($buffer.ToArray())
    }
    $zip.Dispose()
    # Same calls as Ribbon.HighQualityPdf.cs: ppRelativeToSlide with 4x the slide size renders at 384 dpi.
    $renders = New-Object 'System.Collections.Generic.List[byte[]]'
    foreach ($index in 1, 2, 3, 6) {
        $renderPath = Join-Path $outputDirectory "render-$index.png"
        $slide.Shapes.Range($index).Export($renderPath, 2, [int]($presentation.PageSetup.SlideWidth * 4),
            [int]($presentation.PageSetup.SlideHeight * 4), 1)
        $renders.Add([IO.File]::ReadAllBytes($renderPath))
    }

    $svg = [Xml.Linq.XDocument]::Load($svgPath)
    $total = [SlideSCI.InkscapePdfExporter]::CountEmbeddedImages($svg)
    $replaced = [SlideSCI.InkscapePdfExporter]::ReplaceEmbeddedImages($svg, $images, $renders)
    $svg.Save($svgPath, [Xml.Linq.SaveOptions]::DisableFormatting)
    # The brightened and perspective pictures differ from their originals, so they are re-rendered instead.
    if ($total -ne 4 -or $replaced.Originals -ne 2 -or $replaced.Rerendered -ne 2) {
        throw "Expected 2 originals and 2 re-rendered of 4, got $($replaced.Originals) and $($replaced.Rerendered) of $total."
    }

    $pdfPath = Join-Path $outputDirectory 'selection.pdf'
    [SlideSCI.InkscapePdfExporter]::ConvertSvgToPdf($inkscape, $svgPath, $pdfPath, 180000)
    $pdf = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($pdfPath))
    $widths = [regex]::Matches($pdf, '/Width (\d+)') | ForEach-Object { [int]$_.Groups[1].Value }
    if (-not ($widths -contains 2000)) { throw "The PDF does not contain the 2000 px original. Image widths: $($widths -join ', ')" }
    if (-not ($widths -contains 1000)) { throw "The PDF does not contain the 1000 px transparent original. Image widths: $($widths -join ', ')" }
    $pdfToText = Get-Command pdftotext -ErrorAction SilentlyContinue
    if ($pdfToText) {
        $textFile = Join-Path $outputDirectory 'selection.txt'
        & $pdfToText.Source -enc UTF-8 $pdfPath $textFile
        $extracted = [IO.File]::ReadAllText($textFile, [Text.Encoding]::UTF8)
        if (-not $extracted.Contains($caption)) { throw "The PDF text is not searchable as '$caption'." }
    }
    "PASS: $($replaced.Originals) originals and $($replaced.Rerendered) re-rendered of $total; PDF image widths: $($widths -join ', '); $pdfPath"
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($createdApp -and $app.Presentations.Count -eq 0) { $app.Quit() }
}
