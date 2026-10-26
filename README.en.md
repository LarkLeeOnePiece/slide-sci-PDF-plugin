[简体中文](README.md) | English

# SlideSCI HD PDF Edition

A PowerPoint add-in for scientific figures and research presentations, modified from [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) v2.2.4. It keeps every feature of the original add-in and adds **HD PDF export**: selected pictures and shapes are exported to a PDF in which pictures keep their original resolution while text and shapes stay vector.

> **Modification notice:** Modified by LarkLeeOnePiece in October 2026 to add the HD PDF export and an installer. For the original author's feature overview and tutorials, see the archived copies: [Original README (Chinese)](docs/upstream/README.md) · [Original README (English)](docs/upstream/README.en.md)

The add-in's buttons, dialogs and installer are in Chinese. This README gives the Chinese labels together with their English meaning so you can find them in the interface.

## Why HD PDF

PowerPoint's built-in "Export to PDF" recompresses pictures as JPEG at a lower resolution (a 2000 px picture ends up at about 200 ppi). Exporting manually via "Save as Picture → SVG → Inkscape → PDF" keeps text as vectors, but the pictures inside the SVG have already been compressed by PowerPoint.

HD PDF runs the whole pipeline in one click and repairs the quality loss along the way:

- Plain pictures are swapped back to the **original images** stored in the presentation and embedded losslessly.
- Pictures with PowerPoint effects (3D rotation, brightness, …) are re-rendered at **4× resolution (about 384 dpi)**.
- Text is searchable with embedded fonts, shapes and lines stay vector, and transparency and cropping are preserved.

## Installation

### Requirements

| Item | Requirement |
|---|---|
| OS | Windows 10 / 11 (Mac is not supported) |
| PowerPoint | Microsoft 365 / PowerPoint 2016 or later |
| HD PDF | [Inkscape](https://inkscape.org/release/) 1.0 or later (tested with 1.4.2; the add-in's other features do not need it) |

.NET Framework 4.7.2 and the [VSTO runtime](https://www.microsoft.com/download/details.aspx?id=105522) are already present on most PCs with Office; the installer checks for them.

### Steps

1. Download `SlideSCI-PDF-Setup-<version>.exe` from [Releases](https://github.com/LarkLeeOnePiece/slide-sci-PDF-plugin/releases).
2. **Close all PowerPoint windows**, then run the installer. It installs for the current user and does not need administrator rights.
3. Follow the prompts. The installer automatically:
   - checks for .NET Framework and the VSTO runtime and offers download links if they are missing;
   - detects the original SlideSCI's add-in registration and offers to remove it (while both are registered, this version's tabs do not appear);
   - detects Inkscape and offers to open its download page if it is not installed.
4. Open PowerPoint. Two ribbon tabs, **SlideSCI** and **SlideSCI(2)**, appear.

The first time you run the installer, Windows SmartScreen may say "Windows protected your PC" because the installer is not signed with a commercial code-signing certificate. Click "More info → Run anyway".

## Using HD PDF

1. Select the pictures or shapes to export on a slide. Multiple selections and whole groups are supported.
2. Click **SlideSCI tab → 导出 (Export) → 高清PDF (HD PDF)**.
3. On first use the add-in looks for Inkscape automatically. If it cannot find it, it asks you to pick `inkscape.exe` (in the `bin` folder of the Inkscape installation) and remembers your choice.
4. Choose where to save the PDF. Conversion usually takes a few seconds, during which PowerPoint briefly stops responding.
5. When it finishes, a summary of how the pictures were handled is shown, and you are asked whether to open the PDF. For example:

   > 共 11 张图片：5 张使用原始图片，6 张带三维旋转等效果的图片按 4 倍分辨率重新渲染。
   > (11 pictures: 5 use the original images, 6 with effects such as 3D rotation were re-rendered at 4× resolution.)

**Notes**

- A multiple selection is exported as one PDF with the objects' relative positions preserved.
- The PDF page has the selection's real size on the slide. The SVG produced by PowerPoint's "Save as Picture" is scaled by the display scaling factor, so the PDF from the manual workflow can be several times larger in page size. The content and proportions are the same; only the overall scale differs.
- Original pictures are embedded losslessly. If the presentation contains pictures far larger than their displayed size (for example a 3840 px picture shown 6 cm wide), the PDF will be large; the extra detail is not visible when viewing or printing normally.

The original add-in's other features (picture captions, sub-figure labels, alignment, copy/paste position, code blocks, LaTeX/Markdown, zoom-in insets, scale bars, pseudo-color, shape library, AI assistant, …) work as before; see the [original README](docs/upstream/README.en.md).

## FAQ

**The tabs do not appear and there is no message**
Usually the original SlideSCI's add-in registration `Achuan-2.SlideSCI` is still in the registry. Run the installer again and choose to remove it when asked, or delete it manually and restart PowerPoint:

```powershell
Remove-Item 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Achuan-2.SlideSCI'
```

You can also check under "File → Options → Add-ins" whether SlideSCI is listed under "Active Application Add-ins".

**`Custom UI Runtime Error in SlideSCI: Unknown Office control ID: ShapeFillColor` pops up**
This only appears when "File → Options → Advanced → General → Show add-in user interface errors" is checked. The built-in icon used by the copy/paste shape style buttons no longer exists in recent Office versions; only those two buttons' icons are affected, and they still work. Click "OK" or uncheck that option.

**Inkscape is not found**
Make sure Inkscape 1.0 or later is installed, and pick `inkscape.exe` in the `bin` folder of the Inkscape installation when asked.

## Uninstall

Close PowerPoint, then uninstall "SlideSCI 高清PDF版（PowerPoint 插件）" in "Settings → Apps".

## How it works

1. PowerPoint exports the selection to SVG. In this step PowerPoint downsamples pictures and converts them to JPEG.
2. The original images stored in the presentation are read and matched by content, and the corresponding pictures in the SVG are swapped back to the originals. Only positions that really show the same image are replaced.
3. Pictures with effects no longer match their originals, so each such shape is exported separately as a 4× resolution PNG, aligned by its opaque area and placed back, keeping exactly the geometry PowerPoint produced.
4. Inkscape converts the SVG to PDF.

The core code is in `SlideSCI/components/InkscapePdfExporter.cs`; the button handler is in `SlideSCI/Ribbon.HighQualityPdf.cs`.

## For developers: building from source

If you installed the add-in with the installer, uninstall it first so the two registrations do not overwrite each other. The original SlideSCI's add-in registration must also be removed (see the [FAQ](#faq)).

### 1. Prerequisites

- Visual Studio 2019 or 2022 with the **Office/SharePoint development** workload
- [Inkscape](https://inkscape.org/release/) 1.0 or later (needed to run HD PDF and the tests)

### 2. Get the code and restore NuGet packages

```powershell
git clone https://github.com/LarkLeeOnePiece/slide-sci-PDF-plugin.git
cd slide-sci-PDF-plugin
nuget restore SlideSCI.sln
```

Without `nuget.exe`, open the solution in Visual Studio and right-click the solution → "Restore NuGet Packages".

### 3. Prepare a signing certificate

The project is configured with the original author's manifest signing certificate, whose private key you do not have, so use your own certificate. Create a self-signed certificate for local use:

```powershell
$cert = New-SelfSignedCertificate -Subject 'CN=SlideSCI Local Dev' -Type CodeSigningCert `
    -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(10)
$cert.Thumbprint   # note this thumbprint; building and packaging need it
```

The assembly strong name uses `SlideSCI/SlideSCI_TemporaryKey.pfx` from the repository (no password).

### 4. Build

**Option A: Visual Studio**

1. Open `SlideSCI.sln`, right-click the SlideSCI project → "Properties" → "Signing" → check "Sign the ClickOnce manifests" → "Select from Store", and choose `SlideSCI Local Dev`.
2. If the first build asks for the password of `SlideSCI_TemporaryKey.pfx`, leave it empty and confirm.
3. Press F5 to build and start PowerPoint.

This changes `ManifestCertificateThumbprint` in `SlideSCI.csproj`; do not commit that line.

**Option B: command line (does not modify the project file)**

```powershell
& 'C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe' `
    SlideSCI.sln /p:Configuration=Debug /p:ManifestCertificateThumbprint=<thumbprint>
```

For VS 2022, MSBuild is at `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`.

If the command-line build fails with `MSB3325: Cannot import the following key file: SlideSCI_TemporaryKey.pfx`, import the strong-name key into the key container named in the error message (it looks like `VS_KEY_XXXXXXXXXXXXXXXX` and differs per machine):

```powershell
$container = 'VS_KEY_XXXXXXXXXXXXXXXX'   # replace with the name from the error message
$pfx = New-Object Security.Cryptography.X509Certificates.X509Certificate2(
    (Resolve-Path 'SlideSCI\SlideSCI_TemporaryKey.pfx').Path, '',
    [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
$csp = New-Object Security.Cryptography.CspParameters(1)
$csp.KeyContainerName = $container
$csp.KeyNumber = 2
$csp.Flags = [Security.Cryptography.CspProviderFlags]::UseMachineKeyStore
$rsa = New-Object Security.Cryptography.RSACryptoServiceProvider($csp)
$rsa.PersistKeyInCsp = $true
$rsa.ImportParameters($pfx.PrivateKey.ExportParameters($true))
```

Then run the build command again.

After a successful build, VSTO registers the add-in for the current user's PowerPoint. If PowerPoint shows "Publisher cannot be verified" on start, click "Install" (it asks only once). Close PowerPoint before rebuilding, otherwise the add-in DLL is locked and the build fails. To unregister, run "Build → Clean Solution" in Visual Studio, or delete the registry key `HKCU\Software\Microsoft\Office\PowerPoint\Addins\SlideSCI`.

### 5. Package the installer

Install [Inno Setup 6](https://jrsoftware.org/isdl.php), then run:

```powershell
powershell -File installer\build.ps1 -Thumbprint <thumbprint>
```

The script builds the Release configuration (without debug symbols) and uses `installer/SlideSCI-PDF.iss` to create `artifacts\installer\SlideSCI-PDF-Setup-<version>.exe`, taking the version from `SlideSCI/Properties/AssemblyInfo.cs`. The installer does not bundle Inkscape; it only checks for it during installation.

Building the Release configuration also registers the add-in to `bin\Release`; running the generated installer points the registration at the installation folder instead.

### 6. Tests

Test scripts are in `tests/`. The end-to-end HD PDF test needs PowerPoint and Inkscape:

```powershell
powershell -STA -File tests\InkscapePdf.PowerPoint.Tests.ps1
```

## Acknowledgements and license

- Original project: [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) by Achuan-2. The original README (feature overview, tutorials, acknowledgements and disclaimer) is archived at [docs/upstream/README.en.md](docs/upstream/README.en.md).
- HD PDF relies on [Inkscape](https://inkscape.org/) to convert SVG to PDF; Inkscape is not distributed with this add-in.
- The installer is built with [Inno Setup](https://jrsoftware.org/isinfo.php).
- This project keeps the original project's [AGPL-3.0](LICENSE) license, and the original project's disclaimer in its README also applies.
