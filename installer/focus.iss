; FOCUS installer for PyInstaller one-dir output (dist\focus)
; Build example:
;   ISCC.exe "/DAppVersion=1.0.0" installer\focus.iss

#define MyAppName "Focus"
#define MyAppPublisher "Focus"
#define MyAppExeName "focus.exe"
#define MyAppId "{{8BAABC11-ED0F-4A29-B2A5-61DABFF0A24A}}"

#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppVerName={#MyAppName} - {#AppVersion}
AppPublisher={#MyAppPublisher}
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

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\focus\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

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

