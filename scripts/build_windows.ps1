[CmdletBinding()]
param(
    [string]$IconPath = "",
    [string]$PythonPath = "",
    [string]$PfxPath = "",
    [string]$PfxPasswordEnv = "ZADOO_SIGN_PFX_PASSWORD",
    [switch]$NoSelfSign,
    [string]$InnoPath = "",
    [string]$NodePath = "",
    [string]$SignToolPath = "",
    [switch]$SkipInstaller,
    [switch]$SkipPortable,
    [switch]$SkipTests,
    [switch]$KeepOneDir,
    [switch]$InstallMissingTools
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$IconPath = if ($IconPath) { $IconPath } else { Join-Path $Root "app_icon.ico" }
$CloudflaredVersion = "2026.6.0"
$CloudflaredSha256 = "03e322598e84d77406fa55b93f59e8e54636c5d8501d9dce36697fcf080ed8cc"
$PythonVersion = "3.11.9"
$NodeVersion = "24.18.0"
$InnoSetupVersion = "7.0.1-beta"
$WindowsSdkVersion = "10.0.26100.7705"
$WindowsSdkBinVersion = "10.0.26100.0"

if ($PfxPath -and $NoSelfSign) {
    throw "Use either -PfxPath or -NoSelfSign, not both."
}
if (-not $PfxPath -and -not $NoSelfSign) {
    throw "Signing choice required: pass -PfxPath for release signing or -NoSelfSign for an unsigned local build."
}

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Assert-X64PE([string]$Path, [string]$Label) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 64) { throw "$Label is not a valid PE file: $Path" }
    $peOffset = [BitConverter]::ToInt32($bytes, 0x3C)
    if ($peOffset -lt 0 -or $peOffset + 6 -gt $bytes.Length) { throw "$Label has an invalid PE header: $Path" }
    if ([BitConverter]::ToUInt16($bytes, $peOffset + 4) -ne 0x8664) {
        throw "$Label must be x64 (PE machine 0x8664): $Path"
    }
}

function Find-Exe([string[]]$Candidates) {
    foreach ($candidate in $Candidates) {
        if (-not $candidate) { continue }
        $cmd = Get-Command $candidate -ErrorAction Ignore
        if ($cmd) { return $cmd.Source }
        if (Test-Path -LiteralPath $candidate) { return (Resolve-Path $candidate).Path }
    }
    return $null
}

function Resolve-Inno {
    if ($InnoPath) {
        if (-not (Test-Path -LiteralPath $InnoPath)) { throw "Inno Setup compiler not found: $InnoPath" }
        $resolved = (Resolve-Path $InnoPath).Path
        Assert-X64PE $resolved "Inno Setup compiler"
        return $resolved
    }
    $candidates = @(
        "ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 7\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 7\ISCC.exe"
    )
    $found = Find-Exe $candidates
    if (-not $found) { throw "Inno Setup 7 x64 ISCC.exe was not found. Install it or pass -InnoPath." }
    Assert-X64PE $found "Inno Setup compiler"
    return $found
}

function Resolve-SignTool {
    if ($SignToolPath) {
        if (-not (Test-Path -LiteralPath $SignToolPath)) { throw "signtool.exe not found: $SignToolPath" }
        $resolved = (Resolve-Path $SignToolPath).Path
        Assert-X64PE $resolved "signtool.exe"
        return $resolved
    }
    $resolved = "$env:ProgramFiles(x86)\Windows Kits\10\bin\$WindowsSdkBinVersion\x64\signtool.exe"
    if (-not (Test-Path -LiteralPath $resolved)) {
        throw "Windows SDK $WindowsSdkVersion x64 signtool.exe was not found at $resolved; install the pinned SDK or pass -SignToolPath."
    }
    Assert-X64PE $resolved "signtool.exe"
    return $resolved
}

function Install-ToolsIfRequested {
    if (-not $InstallMissingTools) { return }
    Write-Step "Installing missing packaging tools with winget"
    $winget = Find-Exe @("winget.exe")
    if (-not $winget) { throw "winget.exe not found; install Inno Setup and the Windows SDK manually." }
    & $winget install --id JRSoftware.InnoSetup.7 -e --version $InnoSetupVersion --source winget --architecture x64 --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw "winget failed to install Inno Setup $InnoSetupVersion (exit $LASTEXITCODE)." }
    & $winget install --id Microsoft.WindowsSDK.10.0 -e --version $WindowsSdkVersion --source winget --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw "winget failed to install Windows SDK $WindowsSdkVersion (exit $LASTEXITCODE)." }
}

function Resolve-Node {
    if ($NodePath) {
        if (-not (Test-Path -LiteralPath $NodePath)) { throw "Node.js executable not found: $NodePath" }
        $found = (Resolve-Path $NodePath).Path
    } else {
        $found = Find-Exe @("node.exe", "$env:ProgramFiles\nodejs\node.exe")
    }
    $localRoot = "$env:LOCALAPPDATA\Programs\nodejs"
    if (-not $found -and (Test-Path -LiteralPath $localRoot)) {
        $candidate = Get-ChildItem -LiteralPath $localRoot -Recurse -Filter node.exe |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($candidate) { $found = $candidate.FullName }
    }
    if (-not $found) { throw "Node.js x64 was not found. Install it or pass -NodePath." }
    Assert-X64PE $found "Node.js executable"
    $version = (& $found --version | Select-Object -First 1).Trim().TrimStart("v")
    if ($version -ne $NodeVersion) {
        throw "Node.js $NodeVersion x64 is required; resolved Node.js $version at $found"
    }
    return $found
}

function Get-Python {
    if ($PythonPath) {
        if (-not (Test-Path -LiteralPath $PythonPath)) { throw "Python executable not found: $PythonPath" }
        $exe = (Resolve-Path $PythonPath).Path
    } else {
        $launcher = Find-Exe @("py.exe")
        if ($launcher) {
            $launcherOutput = & $launcher -3.11-64 -c "import sys; print(sys.executable)" 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "py launcher failed to resolve Python 3.11 x64: $($launcherOutput -join ' ')"
            }
            $exe = $launcherOutput | Select-Object -First 1
        }
    }
    if (-not $exe) {
        foreach ($candidate in @(
            "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
            "$env:ProgramFiles\Python311\python.exe"
        )) {
            if (Test-Path -LiteralPath $candidate) { $exe = $candidate; break }
        }
    }
    if (-not $exe) {
        throw "Python 3.11 x64 was not found. Install it or register it with the py launcher."
    }
    $exe = $exe.Trim()
    $version = (& $exe -c "import platform; print(platform.python_version())" | Select-Object -First 1).Trim()
    if ($version -ne $PythonVersion) {
        throw "Python $PythonVersion is required; resolved Python $version at $exe"
    }
    $bits = (& $exe -c "import struct; print(struct.calcsize('P') * 8)" | Select-Object -First 1).Trim()
    if ($bits -ne "64") {
        throw "Python 3.11 resolved to a $bits-bit interpreter; Zadoo requires x64: $exe"
    }
    return $exe
}

function Ensure-Cloudflared {
    $cloudDir = Join-Path $Root "build\cloudflared\$CloudflaredVersion\x64"
    New-Item -ItemType Directory -Force -Path $cloudDir | Out-Null
    $target = Join-Path $cloudDir "cloudflared.exe"
    if (-not (Test-Path -LiteralPath $target)) {
        $url = "https://github.com/cloudflare/cloudflared/releases/download/$CloudflaredVersion/cloudflared-windows-amd64.exe"
        Write-Host "Downloading $url"
        Invoke-WebRequest -Uri $url -OutFile $target
    }
    $actualHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne $CloudflaredSha256) {
        throw "cloudflared $CloudflaredVersion SHA256 mismatch: expected $CloudflaredSha256, got $actualHash"
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $target
    if ($signature.Status -ne "Valid") {
        throw "cloudflared signature check failed for $target ($($signature.Status))."
    }
    Assert-X64PE $target "cloudflared"
    return $target
}

function Invoke-Sign([string]$PathToSign) {
    if (-not (Test-Path -LiteralPath $PathToSign)) { throw "Sign target not found: $PathToSign" }
    if ($NoSelfSign) {
        Write-Host "Signing skipped for $PathToSign (-NoSelfSign)."
        return
    }
    if (-not (Test-Path -LiteralPath $PfxPath)) { throw "PFX not found: $PfxPath" }
    $password = [Environment]::GetEnvironmentVariable($PfxPasswordEnv)
    if (-not $password) { throw "PFX password env var is not set: $PfxPasswordEnv" }
    $tool = Resolve-SignTool
    & $tool sign /fd SHA256 /td SHA256 /tr https://timestamp.digicert.com /f $PfxPath /p $password $PathToSign
    if ($LASTEXITCODE -ne 0) { throw "signtool failed with exit code $LASTEXITCODE for $PathToSign" }
    $verified = Get-AuthenticodeSignature -LiteralPath $PathToSign
    if ($verified.Status -ne "Valid") {
        throw "Signature verification failed for ${PathToSign}: $($verified.Status) $($verified.StatusMessage)"
    }
}

function Ensure-Venv([string]$PythonExe) {
    $venv = Join-Path $Root ".build_envs\py311-x64"
    $venvPython = Join-Path $venv "Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython)) {
        Write-Step "Creating venv $venv"
        & $PythonExe -m venv $venv | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Failed to create x64 build venv" }
    }
    $venvRuntime = (& $venvPython -c "import platform,struct; print(platform.python_version() + ':' + str(struct.calcsize('P') * 8))" | Select-Object -First 1).Trim()
    if ($venvRuntime -ne "${PythonVersion}:64") {
        throw "Build environment must use Python $PythonVersion x64; found $venvRuntime at $venvPython"
    }
    & $venvPython -m pip install --no-cache-dir --upgrade "pip==26.1.2" "wheel==0.47.0" | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip/wheel" }
    & $venvPython -m pip install --no-cache-dir $Root -r (Join-Path $Root "build_requirements.txt") | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Failed to install requirements" }
    $sanityImports = "import websockets,mss,PIL,numpy,keyboard,av,sounddevice,soundcard,bettercam,imagecodecs,winpty; import win32api,win32clipboard,win32crypt"
    & $venvPython -c $sanityImports | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Dependency import check failed" }
    return $venvPython
}

function Invoke-SourceChecks([string]$PythonExe) {
    if ($SkipTests) { return }
    Write-Step "Running source checks"
    & $PythonExe -m compileall zadoo_vnc scripts
    if ($LASTEXITCODE -ne 0) { throw "Python compilation check failed" }
    & $PythonExe -m ruff check zadoo_vnc scripts
    if ($LASTEXITCODE -ne 0) { throw "Ruff check failed" }
    & $PythonExe -m vulture zadoo_vnc scripts --min-confidence 80
    if ($LASTEXITCODE -ne 0) { throw "Vulture check failed" }
    & $PythonExe scripts\smoke_test.py
    if ($LASTEXITCODE -ne 0) { throw "Smoke test failed" }
    $node = Resolve-Node
    $env:Path = (Split-Path -Parent $node) + [System.IO.Path]::PathSeparator + $env:Path
    & $PythonExe scripts\check_template_js.py
    if ($LASTEXITCODE -ne 0) { throw "Template JavaScript check failed" }
}

function Build-Zadoo {
    Write-Step "Building Zadoo x64"
    if (-not (Test-Path -LiteralPath $IconPath)) { throw "Icon not found: $IconPath" }
    $python = Get-Python
    $Version = (& $python -c "import pathlib,tomllib; print(tomllib.loads(pathlib.Path('pyproject.toml').read_text(encoding='utf-8'))['project']['version'])" | Select-Object -First 1).Trim()
    $venvPython = Ensure-Venv $python
    Invoke-SourceChecks $venvPython
    $cloudflared = Ensure-Cloudflared

    $workRoot = Join-Path $Root "build\pyinstaller\x64"
    $distRoot = Join-Path $Root "dist"
    $commonArgs = @(
        "--noconfirm",
        "--clean",
        "--windowed",
        "--noupx",
        "--icon", $IconPath,
        "--add-data", "$Root\zadoo_vnc\templates;zadoo_vnc\templates",
        "--add-data", "$Root\zadoo_vnc\static;zadoo_vnc\static",
        "--add-data", "$IconPath;.",
        "--add-data", "$Root\brand-header.png;.",
        "--add-data", "$Root\splash.png;.",
        "--add-data", "$Root\trigger-icon.png;.",
        "--add-binary", "$cloudflared;."
    )
    $hiddenImports = @(
        "win32crypt",
        "pythoncom",
        "pywintypes",
        "tkinter",
        "tkinter.ttk",
        "pygrabber.dshow_core",
        "pygrabber.dshow_ids",
        "soundcard",
        "sounddevice",
        "imagecodecs._shared",
        "imagecodecs._shared_cython"
    )
    $hiddenImports += @("av", "winpty", "bettercam")
    $winptyDir = (& $venvPython -c "import pathlib, winpty; print(pathlib.Path(winpty.__file__).parent)" | Select-Object -First 1).Trim()
    foreach ($winptyBinary in @("winpty-agent.exe", "OpenConsole.exe", "conpty.dll", "winpty.dll")) {
        $winptyPath = Join-Path $winptyDir $winptyBinary
        if (Test-Path -LiteralPath $winptyPath) {
            $commonArgs += @("--add-binary", "$winptyPath;winpty")
        } else {
            throw "Required pywinpty runtime file missing: $winptyPath"
        }
    }
    foreach ($hiddenImport in $hiddenImports) {
        $commonArgs += @("--hidden-import", $hiddenImport)
    }
    $commonArgs += (Join-Path $Root "zadoo_vnc\__main__.py")

    $oneDirDist = Join-Path $distRoot "onedir\x64"
    & $venvPython -m PyInstaller @commonArgs --name "Zadoo" --onedir --distpath $oneDirDist --workpath (Join-Path $workRoot "onedir") --specpath (Join-Path $workRoot "spec-onedir")
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller one-folder build failed for x64" }
    $appExe = Join-Path $oneDirDist "Zadoo\Zadoo.exe"
    Assert-X64PE $appExe "Zadoo one-folder executable"
    Invoke-Sign $appExe

    if (-not $SkipPortable) {
        $portableDist = Join-Path $distRoot "portable"
        & $venvPython -m PyInstaller @commonArgs --name "Zadoo-x64" --onefile --distpath $portableDist --workpath (Join-Path $workRoot "onefile") --specpath (Join-Path $workRoot "spec-onefile")
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller portable build failed for x64" }
        $portableExe = Join-Path $portableDist "Zadoo-x64.exe"
        Assert-X64PE $portableExe "Zadoo portable executable"
        Invoke-Sign $portableExe
    }

    if (-not $SkipInstaller) {
        $iscc = Resolve-Inno
        $installerOut = Join-Path $distRoot "installer"
        New-Item -ItemType Directory -Force -Path $installerOut | Out-Null
        $tempInstallerOut = Join-Path ([System.IO.Path]::GetTempPath()) ("zadoo-installer-x64-" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $tempInstallerOut | Out-Null
        try {
            & $iscc "/DSourceDir=$oneDirDist\Zadoo" "/DOutputDir=$tempInstallerOut" "/DAppIcon=$IconPath" "/DAppVersion=$Version" (Join-Path $Root "installer\zadoo.iss")
            if ($LASTEXITCODE -ne 0) { throw "Inno Setup build failed for x64" }
            $builtInstaller = Join-Path $tempInstallerOut "Zadoo-$Version-x64-Setup.exe"
            if (-not (Test-Path -LiteralPath $builtInstaller)) { throw "Inno Setup did not create $builtInstaller" }
            Assert-X64PE $builtInstaller "Zadoo installer"
            Invoke-Sign $builtInstaller
            $finalInstaller = Join-Path $installerOut "Zadoo-$Version-x64-Setup.exe"
            Move-Item -LiteralPath $builtInstaller -Destination $finalInstaller -Force
        } finally {
            Remove-Item -LiteralPath $tempInstallerOut -Recurse -Force
        }

        # The installer and portable executable are self-contained. Remove the loose
        # one-folder app and PyInstaller work directory unless -KeepOneDir is set.
        if (-not $KeepOneDir) {
            foreach ($leftover in @($oneDirDist, $workRoot)) {
                if ($leftover -and (Test-Path -LiteralPath $leftover)) {
                    Remove-Item -LiteralPath $leftover -Recurse -Force
                }
            }
            # Prune now-empty parent dirs so only dist\installer remains.
            foreach ($parent in @((Join-Path $distRoot "onedir"), (Join-Path $distRoot "portable"), (Join-Path $Root "build"))) {
                if ((Test-Path -LiteralPath $parent) -and -not (Get-ChildItem -LiteralPath $parent -Force)) {
                    Remove-Item -LiteralPath $parent -Recurse -Force
                }
            }
        }
    }
}

Install-ToolsIfRequested
Set-Location $Root
Build-Zadoo

Write-Step "Build complete"
Write-Host "Artifacts are under: $(Join-Path $Root 'dist')"
