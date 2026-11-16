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
VersionInfoProductName={#AppName}
VersionInfoVersion={#AppVersion}
VersionInfoTextVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoProductTextVersion={#AppVersion}
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
WizardStyle=modern dark polar includetitlebar
DisableWelcomePage=no
SetupLogging=yes
PrivilegesRequired=admin
SetupArchitecture=x64
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
LicenseFile=..\LICENSE
InfoBeforeFile=NOTICE.txt

[Messages]
WelcomeLabel1=Your desktop. Within reach.
WelcomeLabel2=Welcome to Zadoo — remote access with a clean workspace and every feature free.%n%nSetup installs the Windows app and a startup task for your Windows sign-in. You can change startup behavior in Preferences.%n%nAfter installation, choose an access code, review permissions, and start sharing. You stay in control.
FinishedHeadingLabel=Your workspace is ready.
FinishedLabel=Zadoo is installed. Open Settings to get connected:%n%n1. Choose your private access code.%n2. Review what viewers are allowed to do.%n3. Select Start Zadoo, then share your connection link.%n%nNo subscription. No feature unlocks. Optional developer support is available inside the app.
SelectDirLabel3=Choose where Zadoo will be installed. The default location works for most people.
SelectTasksLabel2=Choose the shortcuts you would like. Zadoo's Windows startup preference can be changed later in Settings.
InfoBeforeLabel=Before connecting, take a moment to understand what you are sharing.

[Files]
Source: "mascot-poses.png"; Flags: dontcopy
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion
Source: "..\NOTICE"; DestDir: "{app}"; DestName: "NOTICE.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Zadoo\Open Zadoo"; Filename: "{app}\Zadoo.exe"; Parameters: "--open"; IconFilename: "{app}\Zadoo.exe"
Name: "{autodesktop}\Open Zadoo"; Filename: "{app}\Zadoo.exe"; Parameters: "--open"; IconFilename: "{app}\Zadoo.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: unchecked

[Run]
Filename: "{cmd}"; Parameters: "/C schtasks /Create /TN ""Zadoo"" /TR ""\""{app}\Zadoo.exe\"""" /SC ONLOGON /RL HIGHEST /F"; Flags: runhidden waituntilterminated
Filename: "{app}\Zadoo.exe"; Parameters: "--settings"; Description: "Open Zadoo Settings — choose access code and permissions"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C schtasks /Delete /TN ""Zadoo"" /F"; Flags: runhidden waituntilterminated; RunOnceId: "DeleteZadooStartupTask"

[Code]
#include "story.iss"

function InitializeUninstall(): Boolean;
begin
  Result := True;
  { Unattended removal must never wait for input or erase user data. }
  if UninstallSilent() then
  begin
    Log('Silent uninstall: retaining Zadoo settings and logs.');
    Exit;
  end;
  if DirExists(ExpandConstant('{localappdata}\Zadoo')) or
     DirExists(ExpandConstant('{commonappdata}\Zadoo')) then
  begin
    if SuppressibleMsgBox('Delete this Windows user''s Zadoo settings and logs, including any legacy ProgramData copy?', mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDNO) = IDYES then
    begin
      DelTree(ExpandConstant('{localappdata}\Zadoo'), True, True, True);
      DelTree(ExpandConstant('{commonappdata}\Zadoo'), True, True, True);
    end;
  end;
end;
