; Inno Setup 6 script - builds ExcelLibrary_Setup_x.y.z.exe
; Run via build.py (passes /DMyAppVersion=...) or: iscc installer.iss

#define MyAppName "会计工作台"
#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif
#define MyAppExe "ExcelLibrary.exe"

[Setup]
AppId={{B5E8A0F2-3C4D-4E6F-9A1B-2C3D4E5F6A7B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=会计工作台
DefaultDirName={code:GetDefaultDir}
UsePreviousAppDir=no
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
UsePreviousGroup=no
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=ExcelLibrary_Setup_{#MyAppVersion}
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExe}
UninstallDisplayName={#MyAppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\ExcelLibrary\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; v1.3.0 及以前的旧名称快捷方式
Type: files; Name: "{autodesktop}\Excel 图书馆.lnk"
Type: filesandordirs; Name: "{autoprograms}\Excel 图书馆"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Registry]
; the app writes this value itself when auto-start is enabled; remove it on uninstall
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "ExcelLibrary"; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/f /im {#MyAppExe}"; Flags: runhidden; RunOnceId: "KillApp"
Filename: "{sys}\schtasks.exe"; Parameters: "/delete /f /tn ExcelLibrary_DailyScan"; Flags: runhidden; RunOnceId: "DelTask"

[Code]
// v1.6 起：文件默认存放在安装目录下（收件箱、文件库），所以默认装到一个容易找到的位置 C:\会计工作台。
// 以前装在隐藏目录 %LOCALAPPDATA%\Programs\ExcelLibrary 的，升级时改到新位置并删除旧程序；
// 用户自己选过其他位置的，升级时继续用原位置。
const
  UninstKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8A0F2-3C4D-4E6F-9A1B-2C3D4E5F6A7B}_is1';

function OldDefaultDir(): String;
begin
  Result := ExpandConstant('{localappdata}') + '\Programs\ExcelLibrary';
end;

function PrevAppPath(): String;
var s: String;
begin
  Result := '';
  if RegQueryStringValue(HKCU, UninstKey, 'Inno Setup: App Path', s) then Result := s
  else if RegQueryStringValue(HKLM, UninstKey, 'Inno Setup: App Path', s) then Result := s;
end;

function GetDefaultDir(Param: String): String;
var prev: String;
begin
  prev := PrevAppPath();
  if (prev = '') or (CompareText(prev, OldDefaultDir()) = 0) then
    Result := ExpandConstant('{sd}') + '\会计工作台'
  else
    Result := prev;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var rc: Integer;
begin
  // 关闭正在运行的程序（包括旧位置的），避免文件被占用
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/f /im ExcelLibrary.exe', '', SW_HIDE, ewWaitUntilTerminated, rc);
  Sleep(800);
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var old: String;
begin
  if CurStep = ssPostInstall then
  begin
    old := OldDefaultDir();
    if (CompareText(old, ExpandConstant('{app}')) <> 0) and FileExists(old + '\ExcelLibrary.exe') then
      DelTree(old, True, True, True);
  end;
end;
