#define AppName "Zadoo"
#define AppPublisher "Zadoo"
#ifndef AppVersion
#error AppVersion define is required
#endif
#ifndef SourceDir
#error SourceDir define is required
#endif
#ifndef OutputDir
#error OutputDir define is required
#endif
#ifndef AppIcon
#error AppIcon define is required
#endif

[Setup]
AppId={{B35D9680-2BB2-48D8-9A45-8E09D7A1F3B4}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\Zadoo
DefaultGroupName=Zadoo
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=Zadoo-{#AppVersion}-x64-Setup
SetupIconFile={#AppIcon}
UninstallDisplayIcon={app}\Zadoo.exe
WizardImageFile=wizard-large.bmp
WizardSmallImageFile=wizard-small.bmp
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
PrivilegesRequired=admin
SetupArchitecture=x64
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
LicenseFile=NOTICE.txt

[Messages]
WelcomeLabel1=Install Zadoo
WelcomeLabel2=This setup will install Zadoo, create the Windows startup entry, and open the local settings window after Finish.
FinishedHeadingLabel=Zadoo is ready
FinishedLabel=Setup has installed Zadoo on this computer. Click Finish to open Settings.

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Zadoo\Open Zadoo"; Filename: "{app}\Zadoo.exe"; Parameters: "--open"; IconFilename: "{app}\Zadoo.exe"
Name: "{autodesktop}\Open Zadoo"; Filename: "{app}\Zadoo.exe"; Parameters: "--open"; IconFilename: "{app}\Zadoo.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: unchecked

[Run]
Filename: "{cmd}"; Parameters: "/C schtasks /Create /TN ""Zadoo"" /TR ""\""{app}\Zadoo.exe\"""" /SC ONLOGON /RL HIGHEST /F"; Flags: runhidden waituntilterminated
Filename: "{app}\Zadoo.exe"; Parameters: "--settings"; Description: "Launch Zadoo setup"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C schtasks /Delete /TN ""Zadoo"" /F"; Flags: runhidden waituntilterminated; RunOnceId: "DeleteZadooStartupTask"

[Code]
function InitializeUninstall(): Boolean;
begin
  Result := True;
  if DirExists(ExpandConstant('{commonappdata}\Zadoo')) then
  begin
    if MsgBox('Delete Zadoo settings and logs from ProgramData?', mbConfirmation, MB_YESNO) = IDYES then
      DelTree(ExpandConstant('{commonappdata}\Zadoo'), True, True, True);
  end;
end;
