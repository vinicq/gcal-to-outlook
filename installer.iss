; Inno Setup script - GCal -> Teams Sync
; Builds a per-user Setup.exe that installs to %LOCALAPPDATA%\GCalSync.
; No administrator rights required. Registers a native uninstaller in
; Windows "Apps & features". Autostart is controlled from inside the app
; (the "Start with Windows" checkbox), so it is NOT configured here.
;
; Compile with:
;   "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" installer.iss

#define MyAppName "GCal Teams Sync"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Vinicius Queiroz"
#define MyAppExeName "GCalSync.exe"

[Setup]
; A stable GUID identifies the app for upgrades and uninstall.
AppId={{B7F4B2E1-9C3A-4D6E-8F2B-1A5C7E9D0F31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Per-user install: no admin prompt, registers uninstall under HKCU.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\GCalSync
DisableDirPage=yes
DefaultGroupName=GCal Teams Sync
DisableProgramGroupPage=yes
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=installer
OutputBaseFilename=GCalSync-Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "GCalSync.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Remove the scheduled task if a previous version created one. Silent if absent.
Filename: "{sys}\schtasks.exe"; Parameters: "/delete /tn ""GCal-Teams-Sync"" /f"; Flags: runhidden; RunOnceId: "DelSyncTask"

[UninstallDelete]
; Remove the autostart launcher dropped in the Startup folder by the app checkbox.
Type: files; Name: "{userstartup}\GCalSync-autorun.vbs"
; Remove user data created at runtime (config, tokens, db, log) and the folder.
Type: filesandordirs; Name: "{app}"
