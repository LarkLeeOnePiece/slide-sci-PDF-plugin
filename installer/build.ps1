# 编译 Release 版并生成安装包 artifacts\installer\SlideSCI-PDF-Setup-<版本>.exe
# 用法：powershell -File installer\build.ps1 -Thumbprint <清单签名证书指纹>
# 证书需在当前用户「个人」证书库中，生成方法见 README。
param(
    [Parameter(Mandatory = $true)][string]$Thumbprint,
    [string]$MSBuild,
    [string]$Iscc
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

if (-not $MSBuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    $MSBuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' |
        Select-Object -First 1
}
if (-not $MSBuild -or -not (Test-Path $MSBuild)) { throw '未找到 MSBuild，请安装 Visual Studio 或通过 -MSBuild 指定路径。' }
if (-not $Iscc) {
    $Iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $Iscc) { throw '未找到 Inno Setup 6 的 ISCC.exe，请从 https://jrsoftware.org/isdl.php 安装或通过 -Iscc 指定路径。' }

$cert = Get-Item "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
if (-not $cert) { throw "当前用户证书库中没有指纹为 $Thumbprint 的证书。" }
# 与 VSTO 写入信任列表的格式一致：<RSAKeyValue><Modulus>…</Modulus><Exponent>…</Exponent></RSAKeyValue>
$publicKey = $cert.PublicKey.Key.ToXmlString($false)

$assemblyInfo = Get-Content (Join-Path $repo 'SlideSCI\Properties\AssemblyInfo.cs') -Raw
$version = [regex]::Match($assemblyInfo, 'AssemblyVersion\("(\d+\.\d+\.\d+)').Groups[1].Value
if (-not $version) { throw '无法从 AssemblyInfo.cs 读取版本号。' }

Write-Host "编译 Release 版 $version ..."
# 注意：VSTO 编译会顺带把加载项注册到 bin\Release；安装生成的安装包后会改为指向安装目录。
# DebugType=none：不生成 PDB，也不在 DLL 中嵌入编译机器上的 PDB 路径。
& $MSBuild (Join-Path $repo 'SlideSCI.sln') /t:Rebuild /p:Configuration=Release /p:DebugType=none "/p:ManifestCertificateThumbprint=$Thumbprint" /m /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'Release 编译失败。' }

Write-Host '生成安装包 ...'
& $Iscc /Q "/DAppVersion=$version" "/DPublicKey=$publicKey" (Join-Path $PSScriptRoot 'SlideSCI-PDF.iss')
if ($LASTEXITCODE -ne 0) { throw 'Inno Setup 编译失败。' }
$output = Join-Path $repo "artifacts\installer\SlideSCI-PDF-Setup-$version.exe"
Write-Host "完成：$output（$([Math]::Round((Get-Item $output).Length / 1MB, 1)) MB）"
