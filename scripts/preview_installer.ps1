# Produce an inert visual test artifact, NEVER a release installer.
# Reuse the actual branding/messages, but remove every installation operation.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$output = Join-Path $root 'build\wizard-preview-cartoon'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$source = [IO.File]::ReadAllText((Join-Path $root 'installer\zadoo.iss'))
$preview = ([regex]::Split($source, '(?m)^\[Files\]\r?$', 2))[0]
if ($preview -match '(?m)^\[(Files|Run|Icons|Tasks|UninstallRun|Code)\]') { throw 'Unsafe preview source sections' }
$preview = $preview -replace '(?m)^PrivilegesRequired=admin\r?$', 'PrivilegesRequired=lowest'
$preview = $preview -replace '(?m)^OutputBaseFilename=.+\r?$', 'OutputBaseFilename=Zadoo-installer-ui-preview'
$preview = $preview -replace '(?m)^DefaultDirName=.+\r?$', 'DefaultDirName={tmp}\Zadoo-inert-preview'
$preview = $preview -replace '(?m)^AppId=.+\r?$', 'AppId=Zadoo-Inert-Visual-Preview'
$preview = $preview -replace '\[Setup\]', "[Setup]`r`nUninstallable=no`r`nCreateAppDir=no"
foreach ($entry in @{
    WizardImageFile = 'installer\wizard-large.bmp'
    WizardSmallImageFile = 'installer\wizard-small.bmp'
    LicenseFile = 'LICENSE'
    InfoBeforeFile = 'installer\NOTICE.txt'
}.GetEnumerator()) {
    $preview = [regex]::Replace($preview, "(?m)^$($entry.Key)=.+\r?$", "$($entry.Key)=$(Join-Path $root $entry.Value)")
}
$preview += "`r`n[Files]`r`nSource: `"$(Join-Path $root 'installer\mascot-poses.png')`"; Flags: dontcopy`r`n"
$preview += "`r`n[Code]`r`n#define ZadooVisualPreview`r`n#include `"$(Join-Path $root 'installer\story.iss')`"`r`n"
$preview += @'

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := CurPageID <> wpReady;
  if not Result then
    MsgBox('This is an inert UI preview. No files, shortcuts or startup tasks can be installed. Close this window when finished reviewing.', mbInformation, MB_OK);
end;

procedure CancelButtonClick(CurPageID: Integer; var Cancel, Confirm: Boolean);
begin
  Cancel := True;
  Confirm := False;
end;
'@
# Mechanical derivation of the preview keeps its layout identical to the real wizard.
$previewPath = Join-Path $output 'preview.iss'
[IO.File]::WriteAllText($previewPath, $preview, [Text.UTF8Encoding]::new($false))
$versionText = [IO.File]::ReadAllText((Join-Path $root 'zadoo_vnc\__init__.py'))
$version = [regex]::Match($versionText, '__version__\s*=\s*"([^"]+)"').Groups[1].Value
if (-not $version) { throw 'Cannot read the Zadoo version' }
$compiler = Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 7\ISCC.exe'
& $compiler /Qp "/DAppVersion=$version" "/DSourceDir=$root" "/DOutputDir=$output" "/DAppIcon=$(Join-Path $root 'app_icon.ico')" $previewPath
if ($LASTEXITCODE -ne 0) { throw "Preview compiler failed: $LASTEXITCODE" }
Write-Output (Join-Path $output 'Zadoo-installer-ui-preview.exe')
