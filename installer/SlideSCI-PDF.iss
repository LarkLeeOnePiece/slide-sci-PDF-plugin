; SlideSCI（高清PDF版）安装脚本。请通过 installer\build.ps1 编译，它会编译 Release 版并传入下面的定义。
; 不内置 Inkscape：安装时检测，未安装则提示下载。

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef PublicKey
  #error 缺少 PublicKey：VSTO 清单签名证书的 RSA 公钥 XML，由 build.ps1 传入。
#endif
#ifndef SourceDir
  #define SourceDir "..\SlideSCI\bin\Release"
#endif

#define RepoUrl "https://github.com/LarkLeeOnePiece/slide-sci-PDF-plugin"
#define AddinKey "Software\Microsoft\Office\PowerPoint\Addins\SlideSCI"
#define UpstreamAddinKey "Software\Microsoft\Office\PowerPoint\Addins\Achuan-2.SlideSCI"
; 固定的信任记录键名，重复安装时覆盖同一条记录。
#define InclusionKey "Software\Microsoft\VSTO\Security\Inclusion\9a31a2e2-e4b0-4641-95e8-f16a3a9c27b6"

[Setup]
AppId={{D55972A6-2057-48BD-8943-6FF3F50294D3}
AppName=SlideSCI 高清PDF版
AppVersion={#AppVersion}
AppVerName=SlideSCI 高清PDF版 {#AppVersion}
AppPublisher=LarkLeeOnePiece
AppPublisherURL={#RepoUrl}
AppSupportURL={#RepoUrl}/issues
AppUpdatesURL={#RepoUrl}/releases
DefaultDirName={localappdata}\Programs\SlideSCI-PDF
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
LicenseFile=..\LICENSE
OutputDir=..\artifacts\installer
OutputBaseFilename=SlideSCI-PDF-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=SlideSCI 高清PDF版（PowerPoint 插件）
CloseApplications=no

[Languages]
Name: "chs"; MessagesFile: "ChineseSimplified.isl"

[Messages]
FinishedLabel=[name] 已安装完成。%n%n请重新启动 PowerPoint，在「SlideSCI」选项卡中使用插件；选中图片后点击「导出 → 高清PDF」即可导出高清 PDF。%n%n本软件依据 AGPL-3.0 许可证发布，基于 Achuan-2/SlideSCI 修改。源代码：%n{#RepoUrl}

[Files]
; 只打包清单中登记的文件，VSTO 加载时会逐个校验哈希。
Source: "{#SourceDir}\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\SlideSCI.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\SlideSCI.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\SlideSCI.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\latex-converter\*"; DestDir: "{app}\latex-converter"; Flags: ignoreversion recursesubdirs
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "{#AddinKey}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "SlideSCI"
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "Description"; ValueData: "SlideSCI 高清PDF版"
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{code:AppUrl}/SlideSCI.vsto|vstolocal"
; 预先信任清单发布者，首次加载时不再弹出「无法验证发布者」。
Root: HKCU; Subkey: "{#InclusionKey}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#InclusionKey}"; ValueType: string; ValueName: "Url"; ValueData: "file:///{code:AppUrl}/SlideSCI.vsto"
Root: HKCU; Subkey: "{#InclusionKey}"; ValueType: string; ValueName: "PublicKey"; ValueData: "{#PublicKey}"

[Code]
const
  InkscapeDownloadUrl = 'https://inkscape.org/release/';
  VstoDownloadUrl = 'https://www.microsoft.com/zh-cn/download/details.aspx?id=105522';
  NetDownloadUrl = 'https://dotnet.microsoft.com/download/dotnet-framework/net472';

function AppUrl(Param: String): String;
begin
  Result := ExpandConstant('{app}');
  StringChangeEx(Result, '\', '/', True);
end;

procedure OpenUrl(Url: String);
var
  ErrorCode: Integer;
begin
  ShellExec('open', Url, '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
end;

function PowerPointRunning: Boolean;
begin
  Result := FindWindowByClassName('PPTFrameClass') <> 0;
end;

function WaitForPowerPointClosed: Boolean;
begin
  Result := True;
  while PowerPointRunning do
    if MsgBox('PowerPoint 正在运行。请先保存并关闭所有 PowerPoint 窗口，然后点击「重试」。',
      mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
end;

function RegKeyExistsAnyView(SubKey: String): Boolean;
begin
  Result := RegKeyExists(HKLM32, SubKey) or (IsWin64 and RegKeyExists(HKLM64, SubKey));
end;

function VstoRuntimeInstalled: Boolean;
begin
  Result := RegKeyExistsAnyView('SOFTWARE\Microsoft\VSTO Runtime Setup\v4R') or
    RegKeyExistsAnyView('SOFTWARE\Microsoft\VSTO Runtime Setup\v4');
end;

function NetFramework472Installed: Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and
    (Release >= 461808);
end;

function AppPathExists(RootKey: Integer): Boolean;
var
  Path: String;
begin
  Result := RegQueryStringValue(RootKey, 'Software\Microsoft\Windows\CurrentVersion\App Paths\inkscape.exe', '', Path) and
    FileExists(RemoveQuotes(Path));
end;

function InkscapeInstalled: Boolean;
begin
  Result := AppPathExists(HKCU) or AppPathExists(HKLM32) or (IsWin64 and AppPathExists(HKLM64)) or
    (IsWin64 and FileExists(ExpandConstant('{commonpf64}\Inkscape\bin\inkscape.exe'))) or
    FileExists(ExpandConstant('{commonpf32}\Inkscape\bin\inkscape.exe')) or
    FileExists(ExpandConstant('{localappdata}\Programs\Inkscape\bin\inkscape.exe'));
end;

function InitializeSetup: Boolean;
begin
  Result := WaitForPowerPointClosed;
  if not Result then Exit;
  if not NetFramework472Installed then
    if MsgBox('未检测到 .NET Framework 4.7.2 或更高版本，插件需要它才能运行。' + #13#10 + #13#10 +
      '是否打开下载页面？（安装会继续）', mbConfirmation, MB_YESNO) = IDYES then
      OpenUrl(NetDownloadUrl);
  if not VstoRuntimeInstalled then
    if MsgBox('未检测到 Visual Studio 2010 Tools for Office Runtime，插件需要它才能在 PowerPoint 中加载。' + #13#10 + #13#10 +
      '是否打开下载页面？（安装会继续）', mbConfirmation, MB_YESNO) = IDYES then
      OpenUrl(VstoDownloadUrl);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    // 原版 SlideSCI 的注册项即使已禁用，也会让本版本的选项卡无法显示。
    if RegKeyExists(HKCU, '{#UpstreamAddinKey}') then
      if MsgBox('检测到原版 SlideSCI 的加载项注册。它会导致本版本的 SlideSCI 选项卡无法显示。' + #13#10 + #13#10 +
        '是否移除原版的加载项注册？（推荐。原版的程序文件不会被删除，也可以之后在「设置 → 应用」中卸载原版）',
        mbConfirmation, MB_YESNO) = IDYES then
        RegDeleteKeyIncludingSubkeys(HKCU, '{#UpstreamAddinKey}');
  end
  else if CurStep = ssPostInstall then
  begin
    if not InkscapeInstalled then
      if MsgBox('未检测到 Inkscape。「高清PDF」功能需要 Inkscape 1.0 或更高版本，插件的其他功能不受影响。' + #13#10 + #13#10 +
        '是否现在打开 Inkscape 下载页面？', mbConfirmation, MB_YESNO) = IDYES then
        OpenUrl(InkscapeDownloadUrl);
  end;
end;

function InitializeUninstall: Boolean;
begin
  Result := WaitForPowerPointClosed;
end;
