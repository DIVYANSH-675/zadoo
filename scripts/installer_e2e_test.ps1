# Real installer lifecycle tests. Deliberately refuses developer PCs/self-hosted runners.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PreviousInstaller,
    [string]$PreviousVersion = '1.1.4'
)

$ErrorActionPreference = 'Stop'
if ($env:GITHUB_ACTIONS -ne 'true' -or $env:RUNNER_ENVIRONMENT -ne 'github-hosted' -or $env:RUNNER_OS -ne 'Windows') {
    throw 'Installer lifecycle tests require a disposable GitHub-hosted Windows runner.'
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Installer lifecycle tests require an elevated disposable runner.'
}
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$manifest = Get-Content -LiteralPath (Join-Path $root 'dist\release-manifest.json') -Raw | ConvertFrom-Json
$version = [string]$manifest.version
$installer = Join-Path $root "dist\installer\Zadoo-$version-x64-Setup.exe"
$payload = Join-Path $root 'dist\onedir\x64\Zadoo'
$python = Join-Path $root '.build_envs\py311-x64\Scripts\python.exe'
$PreviousInstaller = (Resolve-Path -LiteralPath $PreviousInstaller).Path
$tempRoot = [IO.Path]::GetFullPath($env:RUNNER_TEMP).TrimEnd('\')
$testRoot = [IO.Path]::GetFullPath((Join-Path $tempRoot ('zadoo-lifecycle-' + [guid]::NewGuid().ToString('N'))))
$installRoot = Join-Path $testRoot 'Installed app with spaces'
$app = Join-Path $installRoot 'Zadoo.exe'
$uninstaller = Join-Path $installRoot 'unins000.exe'
$logs = Join-Path $root 'dist\installer-tests'
$registry = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B35D9680-2BB2-48D8-9A45-8E09D7A1F3B4}_is1'
$dataRoots = @((Join-Path $env:LOCALAPPDATA 'Zadoo'), (Join-Path $env:ProgramData 'Zadoo'))
$shortcuts = @(
    (Join-Path ([Environment]::GetFolderPath('CommonPrograms')) 'Zadoo\Open Zadoo.lnk'),
    (Join-Path ([Environment]::GetFolderPath('CommonDesktopDirectory')) 'Open Zadoo.lnk')
)

function Assert-TestInstallPath {
    $resolved = [IO.Path]::GetFullPath($installRoot)
    if (-not $resolved.StartsWith($tempRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path -Leaf $testRoot) -notmatch '^zadoo-lifecycle-[0-9a-f]{32}$' -or
        $resolved -ne (Join-Path $testRoot 'Installed app with spaces')) {
        throw 'Installer test destination escaped its unique runner temporary directory.'
    }
    for ($path = $resolved; $path.Length -ge $tempRoot.Length; $path = Split-Path -Parent $path) {
        if ((Test-Path -LiteralPath $path) -and
            ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Installer test refuses a reparse point: $path"
        }
        if ($path -eq $tempRoot) { break }
    }
}

function Invoke-InstallerProcess([string]$Path, [string[]]$Parameters, [string]$Label) {
    Assert-TestInstallPath
    $process = Start-Process -FilePath $Path -ArgumentList $Parameters -WindowStyle Hidden -PassThru
    $null = $process.Handle
    if (-not $process.WaitForExit(180000)) {
        # Only the process tree started by this test is eligible for timeout cleanup.
        & "$env:WINDIR\System32\taskkill.exe" /PID $process.Id /T /F | Out-Null
        throw "$Label timed out after 180 seconds."
    }
    $process.Refresh()
    if ($process.ExitCode -ne 0) { throw "$Label failed with exit code $($process.ExitCode)." }
    Write-Output "$Label`: PASS (exit 0)"
}

function Install-TestBuild([string]$Path, [string]$Label) {
    Invoke-InstallerProcess $Path @('/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART',
        '/NORESTARTAPPLICATIONS', '/NOCLOSEAPPLICATIONS', '/RESTARTEXITCODE=3010', '/TASKS=desktopicon',
        "/DIR=`"$installRoot`"", "/LOG=`"$(Join-Path $logs "$Label.log")`"") $Label
}

function Assert-Installed([string]$ExpectedVersion, [bool]$ComparePayload) {
    $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($app)
    if ($info.ProductVersion.TrimEnd(' ') -cne $ExpectedVersion -or $info.ProductName.TrimEnd(' ') -cne 'Zadoo') {
        throw 'Installed application product metadata differs.'
    }
    $entry = Get-ItemProperty -LiteralPath $registry
    if ($entry.DisplayVersion -ne $ExpectedVersion -or $entry.InstallLocation.TrimEnd('\') -ne $installRoot) {
        throw 'Installed application registry metadata differs.'
    }
    $task = Get-ScheduledTask -TaskName 'Zadoo' -ErrorAction Stop
    if (@($task.Actions).Count -ne 1 -or $task.Actions[0].Execute.Trim('"') -ne $app) {
        throw 'Startup task does not point to the installed application.'
    }
    $shell = New-Object -ComObject WScript.Shell
    foreach ($path in $shortcuts) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Installer shortcut missing: $path" }
        $shortcut = $shell.CreateShortcut($path)
        if ($shortcut.TargetPath -ne $app -or $shortcut.Arguments -ne '--open') { throw 'Installer shortcut target differs.' }
    }
    if ($ComparePayload) {
        foreach ($file in Get-ChildItem -LiteralPath $payload -File -Recurse) {
            $relative = $file.FullName.Substring($payload.Length + 1)
            $installed = Join-Path $installRoot $relative
            if (-not (Test-Path -LiteralPath $installed) -or
                (Get-FileHash -LiteralPath $installed).Hash -ne (Get-FileHash -LiteralPath $file.FullName).Hash) {
                throw "Installed payload differs: $relative"
            }
        }
        foreach ($name in @('LICENSE', 'NOTICE')) {
            if ((Get-FileHash -LiteralPath (Join-Path $installRoot "$name.txt")).Hash -ne
                (Get-FileHash -LiteralPath (Join-Path $root $name)).Hash) { throw "$name is not bundled correctly." }
        }
    }
    Write-Output "INSTALLED_$ExpectedVersion`: PASS (metadata, startup task, shortcuts, payload=$ComparePayload)"
}

function Assert-DataPreserved {
    foreach ($fixture in $fixtures) {
        if (-not (Test-Path -LiteralPath $fixture.path) -or
            (Get-FileHash -LiteralPath $fixture.path).Hash -ne $fixture.hash) {
            throw 'Installer lifecycle changed a settings or logs fixture.'
        }
    }
    Write-Output 'DATA_PRESERVATION: PASS (current-user and legacy ProgramData settings/logs)'
}

function Uninstall-TestBuild([string]$Mode, [bool]$Suppress, [string]$Label) {
    $parameters = @($Mode, '/NORESTART', "/LOG=`"$(Join-Path $logs "$Label.log")`"")
    if ($Suppress) { $parameters += '/SUPPRESSMSGBOXES' }
    Invoke-InstallerProcess $uninstaller $parameters $Label
    if ((Test-Path -LiteralPath $app) -or (Test-Path -LiteralPath $registry) -or
        (Get-ScheduledTask -TaskName 'Zadoo' -ErrorAction SilentlyContinue)) {
        throw 'Uninstall left the application, registry entry or startup task behind.'
    }
    foreach ($path in $shortcuts) {
        if (Test-Path -LiteralPath $path) { throw 'Uninstall left a shortcut behind.' }
    }
    Assert-DataPreserved
}

Assert-TestInstallPath
foreach ($path in @($registry, $testRoot) + $dataRoots + $shortcuts) {
    if (Test-Path -LiteralPath $path) { throw "Installer test requires a pristine runner; preserving existing path: $path" }
}
if (Get-ScheduledTask -TaskName 'Zadoo' -ErrorAction SilentlyContinue) { throw 'Installer test refuses to replace an existing Zadoo task.' }
if (Get-Process -Name Zadoo -ErrorAction SilentlyContinue) { throw 'Installer test refuses to touch an existing Zadoo process.' }
if (-not (Test-Path -LiteralPath $payload)) { throw 'Installer lifecycle testing requires a -KeepOneDir build.' }
$artifact = @($manifest.artifacts | Where-Object file -eq "installer/Zadoo-$version-x64-Setup.exe")
if ($artifact.Count -ne 1 -or (Get-FileHash -LiteralPath $installer).Hash.ToLowerInvariant() -ne $artifact[0].sha256) {
    throw 'Installer does not match the build manifest.'
}
if ([Diagnostics.FileVersionInfo]::GetVersionInfo($PreviousInstaller).ProductVersion.TrimEnd(' ') -ne $PreviousVersion) {
    throw 'Upgrade baseline version mismatch.'
}
New-Item -ItemType Directory -Path $testRoot, $logs -Force | Out-Null
$fixtures = @()
foreach ($folder in $dataRoots) {
    New-Item -ItemType Directory -Path (Join-Path $folder 'logs') -Force | Out-Null
    foreach ($relative in @('config.json', 'logs\lifecycle-fixture.log')) {
        $path = Join-Path $folder $relative
        [IO.File]::WriteAllText($path, '{"synthetic_lifecycle_fixture":"preserve-me"}', [Text.UTF8Encoding]::new($false))
        $fixtures += @{path=$path; hash=(Get-FileHash -LiteralPath $path).Hash}
    }
}

Install-TestBuild $installer 'fresh-install'
Assert-Installed $version $true
Assert-DataPreserved
Install-TestBuild $installer 'same-version-repair'
Assert-Installed $version $true
Assert-DataPreserved
Uninstall-TestBuild '/SILENT' $false 'silent-uninstall-without-suppression'

Install-TestBuild $PreviousInstaller 'install-upgrade-baseline'
Assert-Installed $PreviousVersion $false
Assert-DataPreserved
Install-TestBuild $installer 'upgrade-to-current'
Assert-Installed $version $true
Assert-DataPreserved
& $python (Join-Path $root 'scripts\runtime_e2e_test.py') --exe $app --control-only
if ($LASTEXITCODE -ne 0) { throw 'Installed application HTTP/auth/assets/stop smoke failed.' }
Uninstall-TestBuild '/VERYSILENT' $true 'very-silent-uninstall-with-suppression'

@{version=$version; upgrade_from=$PreviousVersion; fresh_install='pass'; repair='pass'; upgrade='pass';
    installed_payload_hashes='pass'; startup_task_and_shortcuts='pass'; installed_app_control_smoke='pass';
    silent_uninstall='pass'; very_silent_uninstall='pass'; settings_and_logs_preserved='pass';
    scope='Disposable GitHub-hosted Windows runner; interactive animation and browser/media tested separately'} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $logs 'result.json') -Encoding UTF8
Write-Output 'INSTALLER_LIFECYCLE: PASS'
