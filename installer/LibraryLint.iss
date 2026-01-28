; LibraryLint Inno Setup Script
; Builds installer for LibraryLint media library tool

#define MyAppName "LibraryLint"
#define MyAppVersion "5.2.3"
#define MyAppPublisher "Nick Kliatsko"
#define MyAppURL "https://github.com/kliatsko/librarylint"
#define MyAppExeName "Run-LibraryLint.bat"

[Setup]
; Basic installer info
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Output settings
OutputDir=..\dist
OutputBaseFilename=LibraryLint-{#MyAppVersion}-Setup
SetupIconFile=..\assets\icon.ico
; Compression
Compression=lzma2
SolidCompression=yes
; Installer appearance
WizardStyle=modern
; Privileges - try admin first, fall back to user
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
; Uninstall info
UninstallDisplayIcon={app}\assets\icon.ico
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "installdeps"; Description: "Install dependencies via winget (7-Zip, FFmpeg, yt-dlp, MediaInfo)"; GroupDescription: "Dependencies:"; Flags: unchecked

[Files]
; Main script and launcher
Source: "..\LibraryLint.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Run-LibraryLint.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LibraryLint.psd1"; DestDir: "{app}"; Flags: ignoreversion
; Documentation
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\GETTING_STARTED.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
; Modules
Source: "..\modules\*"; DestDir: "{app}\modules"; Flags: ignoreversion recursesubdirs createallsubdirs
; Config example
Source: "..\config\*"; DestDir: "{app}\config"; Flags: ignoreversion recursesubdirs createallsubdirs
; Icon (if exists)
Source: "..\assets\icon.ico"; DestDir: "{app}\assets"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\assets\icon.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\assets\icon.ico"; Tasks: desktopicon

[Run]
; Run dependency installer if selected
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""Start-Process winget -ArgumentList 'install 7zip.7zip --accept-package-agreements --accept-source-agreements' -Wait; Start-Process winget -ArgumentList 'install Gyan.FFmpeg --accept-package-agreements --accept-source-agreements' -Wait; Start-Process winget -ArgumentList 'install yt-dlp.yt-dlp --accept-package-agreements --accept-source-agreements' -Wait; Start-Process winget -ArgumentList 'install MediaArea.MediaInfo.CLI --accept-package-agreements --accept-source-agreements' -Wait"""; StatusMsg: "Installing dependencies..."; Tasks: installdeps; Flags: runhidden
; Option to launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up any generated files in app directory
Type: filesandordirs; Name: "{app}\logs"

[Code]
// Check if PowerShell 5.1+ is available
function IsPowerShellInstalled(): Boolean;
var
  PSVersion: String;
begin
  Result := RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine',
    'PowerShellVersion', PSVersion);
  if Result then
    Result := (CompareStr(PSVersion, '5.1') >= 0);
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not IsPowerShellInstalled() then
  begin
    MsgBox('LibraryLint requires PowerShell 5.1 or later.' + #13#10 +
           'Please install Windows Management Framework 5.1 or upgrade to Windows 10/11.',
           mbError, MB_OK);
    Result := False;
  end;
end;
