简体中文 | [English](README.en.md)

# SlideSCI 高清PDF版

一个面向科研绘图与组会汇报的 PowerPoint 插件，基于 [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) v2.2.4 修改。在保留原插件全部功能的基础上，新增 **「高清PDF」导出**：把选中的图片和形状导出为图片保持原始分辨率、文字和形状保持矢量的 PDF。

> **修改声明：** 2026 年 10 月由 LarkLeeOnePiece 修改，新增「高清PDF」导出功能和安装程序。原作者的功能介绍与教程见附件：[原 README（中文）](docs/upstream/README.md) · [Original README (English)](docs/upstream/README.en.md)

## 为什么需要高清PDF

PowerPoint 自带的「导出为 PDF」会把图片压缩成 JPEG 并降低分辨率（2000 px 的图片导出后只剩约 200 ppi）。手动「另存为图片 → SVG → Inkscape 转 PDF」能保住矢量文字，但 SVG 里的图片同样已被 PowerPoint 压缩。

「高清PDF」一键完成整个流程，并修复其中的画质损失：

- 普通图片换回文稿中保存的**原始图片**，无损嵌入；
- 带三维旋转、亮度调整等效果的图片按 **4 倍分辨率（约 384 dpi）** 重新渲染；
- 文字可搜索、字体嵌入，形状、线条保持矢量，透明度和裁剪效果不变。

## 安装

### 环境要求

| 项目 | 要求 |
|---|---|
| 系统 | Windows 10 / 11（不支持 Mac） |
| PowerPoint | Microsoft 365 / PowerPoint 2016 及以上 |
| 高清PDF | [Inkscape](https://inkscape.org/release/) 1.0 或更高版本（已在 1.4.2 上测试；插件的其他功能不需要它） |

.NET Framework 4.7.2 和 [VSTO 运行库](https://www.microsoft.com/zh-cn/download/details.aspx?id=105522) 在大多数装有 Office 的电脑上已经存在，安装程序会自动检查。

### 安装步骤

1. 从 [Releases](https://github.com/LarkLeeOnePiece/slide-sci-PDF-plugin/releases) 下载 `SlideSCI-PDF-Setup-<版本>.exe`。
2. **关闭所有 PowerPoint 窗口**，双击运行安装包。安装到当前用户目录，不需要管理员权限。
3. 按提示完成安装。安装程序会自动：
   - 检查 .NET Framework 和 VSTO 运行库，缺少时提示下载；
   - 检测原版 SlideSCI 的加载项注册，询问后移除（两者同时存在时，本版本的选项卡不会显示）；
   - 检测 Inkscape，未安装时提示打开下载页面。
4. 打开 PowerPoint，功能区会出现 **SlideSCI** 和 **SlideSCI(2)** 两个选项卡。

首次运行安装包时，Windows SmartScreen 可能提示「Windows 已保护你的电脑」。这是因为安装包没有商业代码签名证书，点「更多信息 → 仍要运行」即可。

## 使用高清PDF

1. 在幻灯片上选中要导出的图片或形状。可以多选，也可以选中整个组合。
2. 点击 **SlideSCI 选项卡 → 导出 → 高清PDF**。
3. 首次使用时插件会自动查找 Inkscape；找不到时会请你手动选择 `inkscape.exe`（位于 Inkscape 安装目录的 `bin` 下），之后会记住。
4. 选择 PDF 的保存位置。转换通常需要几秒，期间 PowerPoint 会短暂无响应。
5. 完成后会显示图片的处理情况，并询问是否打开 PDF，例如：

   > 共 11 张图片：5 张使用原始图片，6 张带三维旋转等效果的图片按 4 倍分辨率重新渲染。

**说明**

- 多选时导出为一个 PDF，对象之间的相对位置保持不变。
- PDF 页面就是所选对象在幻灯片上的真实尺寸。PowerPoint「另存为图片」得到的 SVG 会按屏幕缩放倍率放大，所以手动流程得到的 PDF 页面可能大好几倍；两者内容比例相同，只是整体缩放不同。
- 原始图片按无损方式嵌入。如果文稿里有远超显示尺寸的大图（例如 3840 px 的图只占 6 cm 宽），PDF 会比较大，多出的细节在正常查看和打印时看不出来。

原插件的其他功能（图片标题、子图标签、排列增强、复制粘贴位置、代码块、LaTeX/Markdown、放大图、比例尺、伪彩、素材库、AI 助手等）用法不变，请参考[原 README](docs/upstream/README.md)。

## 常见问题

**选项卡没有出现，也没有任何提示**
通常是注册表里还留着原版的 `Achuan-2.SlideSCI` 加载项注册。重新运行安装包并在提示时选择移除；或手动删除后重启 PowerPoint：

```powershell
Remove-Item 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Achuan-2.SlideSCI'
```

也可以在「文件 → 选项 → 加载项」中确认 SlideSCI 是否在「活动的应用程序加载项」里。

**弹出 `Custom UI Runtime Error in SlideSCI: Unknown Office control ID: ShapeFillColor`**
只有勾选了「文件 → 选项 → 高级 → 常规 → 显示加载项用户界面错误」才会看到。原因是「复制/粘贴形状样式」按钮使用的内置图标在新版 Office 中已不存在，只影响这两个按钮的图标，功能正常。点「OK」或取消勾选该选项即可。

**提示未找到 Inkscape**
确认已安装 Inkscape 1.0 或更高版本，并在弹出的对话框中选择安装目录 `bin` 下的 `inkscape.exe`。

## 卸载

在「设置 → 应用」中找到「SlideSCI 高清PDF版（PowerPoint 插件）」并卸载。卸载前请先关闭 PowerPoint。

## 工作原理

1. 用 PowerPoint 把选中对象导出为 SVG。PowerPoint 在这一步会把图片降采样并转为 JPEG。
2. 读取文稿中保存的原始图片，按内容比对，把 SVG 中对应的图片换回原图。只替换确实是同一张图的位置。
3. 带效果的图片已经和原图不同，改为对该形状单独导出 4 倍分辨率的 PNG，按不透明区域对齐后贴回原位，几何位置与 PowerPoint 的结果一致。
4. 调用 Inkscape 把 SVG 转为 PDF。

核心代码在 `SlideSCI/components/InkscapePdfExporter.cs`，按钮处理在 `SlideSCI/Ribbon.HighQualityPdf.cs`。

## 开发者：从源码编译

如果已经用安装包装过，请先卸载，避免两份注册互相覆盖。同样需要先移除原版 SlideSCI 的加载项注册（见[常见问题](#常见问题)）。

### 1. 准备环境

- Visual Studio 2019 或 2022，安装 **Office/SharePoint 开发** 工作负载
- [Inkscape](https://inkscape.org/release/) 1.0 或更高版本（运行高清PDF 和测试需要）

### 2. 获取代码并还原 NuGet 包

```powershell
git clone https://github.com/LarkLeeOnePiece/slide-sci-PDF-plugin.git
cd slide-sci-PDF-plugin
nuget restore SlideSCI.sln
```

没有 `nuget.exe` 时，用 Visual Studio 打开解决方案，右键解决方案 →「还原 NuGet 包」。

### 3. 准备签名证书

项目默认使用原作者的清单签名证书，本机没有对应私钥，需要换成自己的证书。生成一个本机自用的测试证书：

```powershell
$cert = New-SelfSignedCertificate -Subject 'CN=SlideSCI Local Dev' -Type CodeSigningCert `
    -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(10)
$cert.Thumbprint   # 记下这个指纹，编译和打包时要用
```

程序集强名称使用仓库中的 `SlideSCI/SlideSCI_TemporaryKey.pfx`（无密码）。

### 4. 编译

**方式 A：Visual Studio 界面**

1. 打开 `SlideSCI.sln`，右键 SlideSCI 项目 →「属性」→「签名」→ 勾选「为 ClickOnce 清单签名」→「从存储区选择」，选择 `SlideSCI Local Dev`。
2. 首次编译时如果弹出导入 `SlideSCI_TemporaryKey.pfx` 的密码框，密码留空确认即可。
3. 按 F5 编译并启动 PowerPoint。

这会修改 `SlideSCI.csproj` 中的 `ManifestCertificateThumbprint`，提交代码时不要提交这一行。

**方式 B：命令行（不修改项目文件）**

```powershell
& 'C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe' `
    SlideSCI.sln /p:Configuration=Debug /p:ManifestCertificateThumbprint=<证书指纹>
```

VS 2022 的 MSBuild 路径为 `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`。

命令行编译若报错 `MSB3325: Cannot import the following key file: SlideSCI_TemporaryKey.pfx`，需要把强名称密钥导入错误信息中给出的密钥容器（名称形如 `VS_KEY_XXXXXXXXXXXXXXXX`，每台电脑不同）：

```powershell
$container = 'VS_KEY_XXXXXXXXXXXXXXXX'   # 替换为错误信息中的名称
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

然后重新运行编译命令。

编译成功后 VSTO 会自动把插件注册到当前用户的 PowerPoint。启动 PowerPoint 时如果弹出「无法验证发布者」，点「安装」即可（只会询问一次）。以后重新编译前需要先关闭 PowerPoint，否则插件 DLL 被占用会导致编译失败。取消注册：在 Visual Studio 中点击「生成 → 清理解决方案」，或删除注册表项 `HKCU\Software\Microsoft\Office\PowerPoint\Addins\SlideSCI`。

### 5. 打包安装程序

安装 [Inno Setup 6](https://jrsoftware.org/isdl.php) 后运行：

```powershell
powershell -File installer\build.ps1 -Thumbprint <证书指纹>
```

脚本会编译 Release 版（不生成调试符号），并用 `installer/SlideSCI-PDF.iss` 生成 `artifacts\installer\SlideSCI-PDF-Setup-<版本>.exe`，版本号取自 `SlideSCI/Properties/AssemblyInfo.cs`。安装包不内置 Inkscape，只在安装时检测。

编译 Release 版时 VSTO 会顺带把加载项注册到 `bin\Release`；运行生成的安装包后会改为指向安装目录。

### 6. 测试

测试脚本位于 `tests/`。高清PDF 的端到端测试需要 PowerPoint 和 Inkscape：

```powershell
powershell -STA -File tests\InkscapePdf.PowerPoint.Tests.ps1
```

## 致谢与许可

- 原项目：[Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)，作者 Achuan-2。原 README（含功能介绍、教程、致谢与免责声明）存档于 [docs/upstream/README.md](docs/upstream/README.md)。
- 高清PDF 依赖 [Inkscape](https://inkscape.org/) 完成 SVG 到 PDF 的转换；Inkscape 不随本插件分发。
- 安装程序使用 [Inno Setup](https://jrsoftware.org/isinfo.php) 制作。
- 本项目沿用原项目的 [AGPL-3.0](LICENSE) 许可证，原项目的[特别说明](docs/upstream/README.md#-特别说明)同样适用。
