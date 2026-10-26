# SlideSCI（高清 PDF 导出版）

本项目基于 [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) v2.2.4 修改，是一个面向科研绘图与组会汇报的 PowerPoint 插件（VSTO，C#）。在保留原插件全部功能的基础上，新增了**「高清PDF」导出**：把选中的图片和形状借助 [Inkscape](https://inkscape.org/) 导出为保留原始图片分辨率的矢量 PDF。

> 原作者的功能介绍、使用教程和视频请见附件：[原 README（中文）](docs/upstream/README.md) · [Original README (English)](docs/upstream/README.en.md)

## 与原版的区别：高清PDF

PowerPoint 自带的「导出为 PDF」会把图片降采样并转为 JPEG（例如 2000 px 的 PNG 会变成约 200 ppi 的 JPEG）。手动「另存为图片 → SVG → Inkscape 转 PDF」能保住矢量和文字，但 SVG 中的图片同样被 PowerPoint 压缩过。

「高清PDF」按钮自动完成整个流程，并在中间修复图片质量：

1. 用 PowerPoint 把选中对象导出为 SVG，版式、旋转、裁剪、透明度和文字都保持矢量。
2. 把 SVG 中被压缩的图片换回文稿里保存的**原始图片**（按内容比对，只替换确实是同一张图的位置）。
3. 带三维旋转、亮度调整等 PowerPoint 效果的图片无法使用原图，改为对该形状单独做 **4 倍分辨率（约 384 dpi）重新渲染**，再按原位置精确贴回。
4. 调用 Inkscape 输出 PDF：文字可搜索、字体嵌入，形状保持矢量。

导出的 PDF 页面就是所选对象在幻灯片上的真实尺寸。PowerPoint 界面里「另存为图片」得到的 SVG 会按屏幕缩放倍率放大，所以手动流程得到的 PDF 页面会大好几倍；两者内容比例相同，只是整体缩放不同。

## 环境要求

| 项目 | 要求 |
|---|---|
| 系统 | Windows 10 / 11（插件不支持 Mac） |
| PowerPoint | Microsoft 365 / PowerPoint 2016 及以上 |
| 运行库 | .NET Framework 4.7.2、[Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/zh-cn/download/details.aspx?id=105522) |
| 高清PDF | [Inkscape](https://inkscape.org/release/) 1.0 或更高版本（已在 1.4.2 上测试） |
| 编译 | Visual Studio 2019 或 2022，安装 **Office/SharePoint 开发** 工作负载 |

## 安装（从源码编译）

本版本没有安装包，需要在本机编译。编译成功后，Visual Studio 会自动把插件注册到当前用户的 PowerPoint，之后每次启动 PowerPoint 都会加载。

### 1. 卸载或移除原版 SlideSCI（重要）

如果装过原作者发布的 SlideSCI，请先在「设置 → 应用」中卸载它。原版在注册表中的条目 `HKCU\Software\Microsoft\Office\PowerPoint\Addins\Achuan-2.SlideSCI` 即使处于禁用状态，也会导致本版本的选项卡**不显示且没有任何提示**。卸载后如果条目还在，可以手动删除：

```powershell
Remove-Item 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Achuan-2.SlideSCI'
```

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
$cert.Thumbprint   # 记下这个指纹，编译时要用
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
    SlideSCI.sln /p:Configuration=Debug /p:ManifestCertificateThumbprint=<上一步的指纹>
```

VS 2022 的 MSBuild 路径为 `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`。

命令行编译若报错 `MSB3325: Cannot import the following key file: SlideSCI_TemporaryKey.pfx`，需要把强名称密钥导入错误信息里给出的密钥容器（名称形如 `VS_KEY_XXXXXXXXXXXXXXXX`，每台电脑不同）：

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

### 5. 启动 PowerPoint

编译成功后启动 PowerPoint，功能区会出现 **SlideSCI** 和 **SlideSCI(2)** 两个选项卡。首次加载时如果弹出「无法验证发布者」，点「安装」即可（测试证书是自签名的，只会询问一次）。

以后修改代码重新编译前，需要先关闭 PowerPoint，否则插件 DLL 被占用会导致编译失败。

## 使用高清PDF

1. 在幻灯片上选中要导出的图片或形状，可以多选，也可以选中组合。
2. 点击 **SlideSCI 选项卡 → 导出 → 高清PDF**。
3. 首次使用会自动查找 Inkscape（依次检查已保存的路径、注册表、PATH 和常见安装目录）；找不到时会让你手动选择 `inkscape.exe`，之后会记住。
4. 选择保存位置。导出期间 PowerPoint 会短暂无响应，转换通常需要几秒。
5. 完成后会提示各类图片的处理情况，例如：

   > 共 11 张图片：5 张使用原始图片，6 张带三维旋转等效果的图片按 4 倍分辨率重新渲染。

多选时导出为一个 PDF，对象之间的相对位置保持不变。

**关于文件大小：** 原始图片按无损方式嵌入。如果文稿里有远超显示尺寸的大图（例如 3840 px 的图只占 6 cm 宽），PDF 会明显比手动流程大，这部分细节在正常查看和打印时看不出来。

原插件的其他功能（图片标题、子图标签、排列、复制粘贴位置、代码块、LaTeX/Markdown、放大图、比例尺、伪彩、素材库、AI 助手等）用法不变，请参考[原 README](docs/upstream/README.md)。

## 常见问题

**选项卡没有出现，也没有任何提示**
几乎都是因为注册表里还留着原版的 `Achuan-2.SlideSCI` 条目，按[安装第 1 步](#1-卸载或移除原版-slidesci重要)删除后重启 PowerPoint。也可以在「文件 → 选项 → 加载项」确认 SlideSCI 是否在「活动的应用程序加载项」中。

**弹出 `Custom UI Runtime Error in SlideSCI: Unknown Office control ID: ShapeFillColor`**
只有勾选了「文件 → 选项 → 高级 → 常规 → 显示加载项用户界面错误」才会看到。原因是「复制/粘贴形状样式」按钮使用的内置图标在新版 Office 中已不存在，只影响这两个按钮的图标，功能正常。点「OK」或取消勾选上述选项即可。

**提示未找到 Inkscape**
确认已安装 Inkscape 1.0 以上版本，并在弹出的对话框里选择安装目录 `bin` 下的 `inkscape.exe`。

**编译时提示文件被占用**
PowerPoint 正在加载插件 DLL，关闭所有 PowerPoint 窗口后再编译。

## 卸载

在 Visual Studio 中点击「生成 → 清理解决方案」即可取消注册；或直接删除注册表条目：

```powershell
Remove-Item 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\SlideSCI'
```

## 开发与测试

- 高清PDF 的核心逻辑在 `SlideSCI/components/InkscapePdfExporter.cs`，按钮处理在 `SlideSCI/Ribbon.HighQualityPdf.cs`。
- 测试脚本位于 `tests/`。高清PDF 的端到端测试需要 PowerPoint 和 Inkscape：

  ```powershell
  powershell -STA -File tests\InkscapePdf.PowerPoint.Tests.ps1
  ```

## 致谢与许可

- 原项目：[Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)，作者 Achuan-2。原 README（含功能介绍、教程、致谢与免责声明）存档于 [docs/upstream/README.md](docs/upstream/README.md)。
- 高清PDF 依赖 [Inkscape](https://inkscape.org/) 完成 SVG 到 PDF 的转换。
- 本项目沿用原项目的 [AGPL-3.0](LICENSE) 许可证，原项目的[特别说明](docs/upstream/README.md#-特别说明)同样适用。
