; FOCUS installer for PyInstaller one-dir output (dist\focus)
; Build example:
;   ISCC.exe "/DWorkbookPassword=<secret>" "/DAppVersion=1.0.0" installer\focus.iss

#define MyAppName "Focus"
#define MyAppPublisher "Focus"
#define MyAppExeName "focus.exe"
#define MyAppId "{{8BAABC11-ED0F-4A29-B2A5-61DABFF0A24A}}"

#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

#ifndef WorkbookPassword
  #error "WorkbookPassword define is required. Pass /DWorkbookPassword=<secret>."
#endif

#if Len(WorkbookPassword) < 12
  #error "WorkbookPassword must be at least 12 characters."
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
OutputDir=dist
OutputBaseFilename=focus-installer
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=..\assets\kare-logo.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\focus\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: string; ValueName: "FOCUS_WORKBOOK_PASSWORD"; ValueData: "{#WorkbookPassword}"; Flags: uninsdeletevalue; Check: IsAdminInstallMode
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "FOCUS_WORKBOOK_PASSWORD"; ValueData: "{#WorkbookPassword}"; Flags: uninsdeletevalue; Check: not IsAdminInstallMode

[Code]
const
  WM_SETTINGCHANGE = $001A;
  SMTO_ABORTIFHUNG = $0002;

function SendMessageTimeout(
  hWnd: LongWord;
  Msg: LongWord;
  wParam: LongWord;
  lParam: string;
  fuFlags: LongWord;
  uTimeout: LongWord;
  var lpdwResult: LongWord
): LongWord;
  external 'SendMessageTimeoutW@user32.dll stdcall';

procedure RefreshEnvironment;
var
  ResultCode: LongWord;
begin
  SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 'Environment', SMTO_ABORTIFHUNG, 5000, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RefreshEnvironment;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    RefreshEnvironment;
end;
