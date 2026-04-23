; FOCUS installer for PyInstaller one-dir output (dist\focus)
; Build example:
;   ISCC.exe "/DAppVersion=1.0.0" installer\focus.iss

#define MyAppName "Focus"
#define MyAppPublisher "Focus"
#define MyAppCopyright "Copyright (c) 2026"
#define MyAppExeName "focus.exe"
#define MyAppId "{{8BAABC11-ED0F-4A29-B2A5-61DABFF0A24A}}"

#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef AppFileVersion
  #define AppFileVersion AppVersion + ".0"
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppVerName={#MyAppName} - {#AppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog commandline
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
OutputDir=dist
OutputBaseFilename=focus-installer
UninstallDisplayName={#MyAppName} - {#AppVersion}
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=..\assets\kare-logo.ico
VersionInfoVersion={#AppFileVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoTextVersion={#AppFileVersion}
VersionInfoProductName={#MyAppName}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Setup
VersionInfoCopyright={#MyAppCopyright}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\focus\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Prevent mixed-runtime upgrades: clear prior bundled runtime before copying.
Type: filesandordirs; Name: "{app}\_internal"

[Dirs]
; Shared workbook secret store (machine scope)
Name: "{commonappdata}\FOCUS\secrets"; Permissions: users-modify

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove shared machine-wide secret store
Type: filesandordirs; Name: "{commonappdata}\FOCUS"
Type: filesandordirs; Name: "{commonappdata}\Focus"

[Code]
function ExtractNumericPart(const Token: string): Integer;
var
  I: Integer;
  Numeric: string;
begin
  Numeric := '';
  for I := 1 to Length(Token) do
  begin
    if (Token[I] >= '0') and (Token[I] <= '9') then
      Numeric := Numeric + Token[I]
    else
      Break;
  end;
  if Numeric = '' then
    Result := 0
  else
    Result := StrToIntDef(Numeric, 0);
end;

function NextVersionPart(var Version: string): Integer;
var
  DotPos: Integer;
  Token: string;
begin
  DotPos := Pos('.', Version);
  if DotPos = 0 then
  begin
    Token := Version;
    Version := '';
  end
  else
  begin
    Token := Copy(Version, 1, DotPos - 1);
    Delete(Version, 1, DotPos);
  end;
  Result := ExtractNumericPart(Token);
end;

function CompareVersionStrings(const A, B: string): Integer;
var
  LeftVersion, RightVersion: string;
  LeftPart, RightPart: Integer;
begin
  LeftVersion := Trim(A);
  RightVersion := Trim(B);
  while (LeftVersion <> '') or (RightVersion <> '') do
  begin
    LeftPart := NextVersionPart(LeftVersion);
    RightPart := NextVersionPart(RightVersion);
    if LeftPart < RightPart then
    begin
      Result := -1;
      Exit;
    end;
    if LeftPart > RightPart then
    begin
      Result := 1;
      Exit;
    end;
  end;
  Result := 0;
end;

function TryGetInstalledVersion(var InstalledVersion: string): Boolean;
var
  UninstallKey: string;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppId}_is1';
  Result :=
    RegQueryStringValue(HKLM64, UninstallKey, 'DisplayVersion', InstalledVersion) or
    RegQueryStringValue(HKLM, UninstallKey, 'DisplayVersion', InstalledVersion) or
    RegQueryStringValue(HKCU, UninstallKey, 'DisplayVersion', InstalledVersion);
end;

function InitializeSetup(): Boolean;
var
  InstalledVersion: string;
  CurrentVersion: string;
begin
  CurrentVersion := '{#AppVersion}';
  if TryGetInstalledVersion(InstalledVersion) then
  begin
    if CompareVersionStrings(CurrentVersion, InstalledVersion) < 0 then
    begin
      MsgBox(
        '{#MyAppName} version ' + InstalledVersion +
        ' is already installed. Downgrade to ' + CurrentVersion +
        ' is not allowed.',
        mbCriticalError,
        MB_OK
      );
      Result := False;
      Exit;
    end;
  end;
  Result := True;
end;

procedure DeleteFocusRoamingDir(const Dir: string);
begin
  if DirExists(Dir) then
  begin
    DelTree(Dir, True, True, True);
  end;
end;

procedure DeleteFocusForAllUsers;
var
  UsersRoot: string;
  FindRec: TFindRec;
  UserDir: string;
begin
  UsersRoot := ExpandConstant('{sd}\Users');
  if not DirExists(UsersRoot) then
    Exit;

  if FindFirst(UsersRoot + '\*', FindRec) then
  try
    repeat
      if ((FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and
         (FindRec.Name <> '.') and (FindRec.Name <> '..') then
      begin
        if (CompareText(FindRec.Name, 'Public') <> 0) and
           (CompareText(FindRec.Name, 'Default') <> 0) and
           (CompareText(FindRec.Name, 'Default User') <> 0) and
           (CompareText(FindRec.Name, 'All Users') <> 0) then
        begin
          UserDir := UsersRoot + '\' + FindRec.Name;
          DeleteFocusRoamingDir(UserDir + '\AppData\Roaming\FOCUS');
          DeleteFocusRoamingDir(UserDir + '\AppData\Roaming\Focus');
        end;
      end;
    until not FindNext(FindRec);
  finally
    FindClose(FindRec);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    DeleteFocusForAllUsers;
  end;
end;

