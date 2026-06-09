[CmdletBinding()]
param(
    [ValidateSet("x64", "x86", "All")]
    [string]$Arch = "x64",
    [string]$IconPath = "C:\Users\divya\Real\app_icon.ico",
    [string]$PfxPath = "",
    [string]$PfxPasswordEnv = "ZADOO_SIGN_PFX_PASSWORD",
    [switch]$NoSelfSign,
    [string]$InnoPath = "",
    [string]$SignToolPath = "",
    [switch]$SkipInstaller,
    [switch]$SkipPortable,
    [switch]$SkipTests,
    [switch]$KeepOneDir,
    [switch]$InstallMissingTools
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$Version = "1.0.0"

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Find-Exe([string[]]$Candidates) {
    foreach ($candidate in $Candidates) {
        if (-not $candidate) { continue }
        try {
            $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
            if ($cmd) { return $cmd.Source }
        } catch {}
        if (Test-Path -LiteralPath $candidate) { return (Resolve-Path $candidate).Path }
    }
    return $null
}

function Resolve-Inno {
    if ($InnoPath) {
        if (-not (Test-Path -LiteralPath $InnoPath)) { throw "Inno Setup compiler not found: $InnoPath" }
        return (Resolve-Path $InnoPath).Path
    }
    $candidates = @(
        "ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $found = Find-Exe $candidates
    if (-not $found) { throw "Inno Setup ISCC.exe was not found. Install Inno Setup 6 or pass -InnoPath." }
    return $found
}

function Resolve-SignTool {
    if ($SignToolPath) {
        if (-not (Test-Path -LiteralPath $SignToolPath)) { throw "signtool.exe not found: $SignToolPath" }
        return (Resolve-Path $SignToolPath).Path
    }
    $found = Find-Exe @("signtool.exe")
    if ($found) { return $found }
    $kitRoot = "$env:ProgramFiles(x86)\Windows Kits\10\bin"
    if (Test-Path -LiteralPath $kitRoot) {
        $candidates = Get-ChildItem -Path $kitRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } |
            Sort-Object FullName -Descending
        if ($candidates) { return $candidates[0].FullName }
    }
    throw "signtool.exe was not found. Install Windows SDK or pass -SignToolPath."
}

function Add-LocalBuildCertTrust([object]$Cert) {
    foreach ($storeName in @("Root", "TrustedPublisher")) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, "CurrentUser")
        $store.Open("ReadWrite")
        try {
            $existing = $store.Certificates.Find(
                [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
                $Cert.Thumbprint,
                $false
            )
            if ($existing.Count -eq 0) {
                $store.Add($Cert)
            }
        } finally {
            $store.Close()
        }
    }
}

function Install-ToolsIfRequested {
    if (-not $InstallMissingTools) { return }
    Write-Step "Installing missing packaging tools with winget"
    $winget = Find-Exe @("winget.exe")
    if (-not $winget) { throw "winget.exe not found; install Inno Setup, Windows SDK, and Python 3.11 x86 manually." }
    & $winget install --id JRSoftware.InnoSetup -e --accept-package-agreements --accept-source-agreements
    & $winget install --id Microsoft.WindowsSDK.10.0 -e --accept-package-agreements --accept-source-agreements
    Write-Host "For x86 builds, install Python 3.11 32-bit if 'py -3.11-32' is not available."
}

function Get-Python([string]$TargetArch) {
    $selector = if ($TargetArch -eq "x64") { "-3.11-64" } else { "-3.11-32" }
    $exe = ""
    try {
        $exe = (& py $selector -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1)
    } catch {}
    if (-not $exe -and $TargetArch -eq "x64") {
        try { $exe = (& py -3.11 -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1) } catch {}
    }
    if (-not $exe -and $TargetArch -eq "x86") {
        $fallbacks = @(
            "$env:LOCALAPPDATA\Programs\Python\Python311-32\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python311-32bit\python.exe",
            "$env:ProgramFiles(x86)\Python311-32\python.exe",
            "$env:ProgramFiles(x86)\Python311\python.exe"
        )
        foreach ($candidate in $fallbacks) {
            if (Test-Path -LiteralPath $candidate) {
                $exe = $candidate
                break
            }
        }
    }
    if (-not $exe) {
        throw "Python 3.11 $TargetArch was not found. Install it and ensure the py launcher supports $selector."
    }
    $exe = $exe.Trim()
    $bits = (& $exe -c "import struct; print(struct.calcsize('P') * 8)" | Select-Object -First 1).Trim()
    $expected = if ($TargetArch -eq "x64") { "64" } else { "32" }
    if ($bits -ne $expected) {
        throw "Python selector for $TargetArch resolved to a $bits-bit interpreter: $exe"
    }
    return $exe
}

function Ensure-Cloudflared([string]$TargetArch) {
    $cloudDir = Join-Path $Root "build\cloudflared\$TargetArch"
    New-Item -ItemType Directory -Force -Path $cloudDir | Out-Null
    $target = Join-Path $cloudDir "cloudflared.exe"
    if (-not (Test-Path -LiteralPath $target)) {
        if ($TargetArch -eq "x64" -and (Test-Path -LiteralPath (Join-Path $Root "cloudflared.exe"))) {
            Copy-Item -LiteralPath (Join-Path $Root "cloudflared.exe") -Destination $target -Force
        } else {
            $suffix = if ($TargetArch -eq "x86") { "386" } else { "amd64" }
            $url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-$suffix.exe"
            Write-Host "Downloading $url"
            Invoke-WebRequest -Uri $url -OutFile $target
        }
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $target
    if ($signature.Status -ne "Valid") {
        throw "cloudflared signature check failed for $target ($($signature.Status))."
    }
    return $target
}

function Invoke-Sign([string]$PathToSign) {
    if (-not (Test-Path -LiteralPath $PathToSign)) { throw "Sign target not found: $PathToSign" }
    if ($PfxPath) {
        if (-not (Test-Path -LiteralPath $PfxPath)) { throw "PFX not found: $PfxPath" }
        $password = [Environment]::GetEnvironmentVariable($PfxPasswordEnv)
        if (-not $password) { throw "PFX password env var is not set: $PfxPasswordEnv" }
        try {
            $tool = Resolve-SignTool
            & $tool sign /fd SHA256 /td SHA256 /tr http://timestamp.digicert.com /f $PfxPath /p $password $PathToSign
            if ($LASTEXITCODE -eq 0) { return }
            throw "signtool failed with exit code $LASTEXITCODE"
        } catch {
            Write-Host "signtool unavailable/failed; signing with PowerShell certificate APIs."
            $secure = ConvertTo-SecureString $password -AsPlainText -Force
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $secure, "Exportable,PersistKeySet")
            $result = Set-AuthenticodeSignature -FilePath $PathToSign -Certificate $cert -HashAlgorithm SHA256
            if ($result.Status -notin @("Valid", "UnknownError")) { throw "PowerShell signing failed for ${PathToSign}: $($result.Status) $($result.StatusMessage)" }
            return
        }
    }
    if ($NoSelfSign) {
        Write-Host "Signing skipped for $PathToSign (no -PfxPath supplied and -NoSelfSign set)."
        return
    }
    $subject = "CN=Zadoo Local Build"
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } | Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $cert) {
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -KeyUsage DigitalSignature -NotAfter (Get-Date).AddYears(5)
    }
    Add-LocalBuildCertTrust $cert
    $result = Set-AuthenticodeSignature -FilePath $PathToSign -Certificate $cert -HashAlgorithm SHA256
    $verified = Get-AuthenticodeSignature -LiteralPath $PathToSign
    if ($result.Status -ne "Valid" -or $verified.Status -ne "Valid") {
        throw "Self-signing failed for ${PathToSign}: $($verified.Status) $($verified.StatusMessage)"
    }
    Write-Host "Signed $PathToSign with local certificate $($cert.Thumbprint)."
}

function Ensure-Venv([string]$TargetArch, [string]$PythonExe) {
    $venv = Join-Path $Root ".build_envs\py311-$TargetArch"
    $venvPython = Join-Path $venv "Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython)) {
        Write-Step "Creating venv $venv"
        & $PythonExe -m venv $venv | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Failed to create venv for $TargetArch" }
    }
    $runtimeRequirements = Join-Path $Root "requirements.txt"
    if ($TargetArch -eq "x86") {
        $x86Requirements = Join-Path $Root "requirements-x86.txt"
        if (Test-Path -LiteralPath $x86Requirements) {
            $runtimeRequirements = $x86Requirements
        }
    }
    & $venvPython -m pip install --no-cache-dir --upgrade pip wheel | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip/wheel for $TargetArch" }
    & $venvPython -m pip install --no-cache-dir -r $runtimeRequirements -r (Join-Path $Root "build_requirements.txt") | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Failed to install requirements for $TargetArch" }
    $sanityImports = "import websockets,mss,pyautogui,PIL,numpy,pyperclip,keyboard,cv2,sounddevice,soundcard,paramiko,imagecodecs; import win32api,win32crypt"
    if ($TargetArch -eq "x64") {
        $sanityImports = "import websockets,mss,pyautogui,PIL,numpy,pyperclip,keyboard,av,cv2,sounddevice,soundcard,paramiko,dxcam,bettercam,imagecodecs,winpty; import win32api,win32crypt"
    }
    & $venvPython -c $sanityImports | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Dependency import check failed for $TargetArch" }
    return $venvPython
}

function Invoke-SourceChecks {
    if ($SkipTests) { return }
    Write-Step "Running source checks"
    & python -m compileall zadoo_vnc zadoo_vnc_single.py scripts
    & python scripts\smoke_test.py
    if (Get-Command node -ErrorAction SilentlyContinue) {
        & python scripts\check_template_js.py
    } else {
        Write-Host "Node.js not found; skipping template JavaScript syntax check."
    }
}

function Build-Arch([string]$TargetArch) {
    Write-Step "Building Zadoo $TargetArch"
    if (-not (Test-Path -LiteralPath $IconPath)) { throw "Icon not found: $IconPath" }
    $python = Get-Python $TargetArch
    $venvPython = Ensure-Venv $TargetArch $python
    $cloudflared = Ensure-Cloudflared $TargetArch

    $workRoot = Join-Path $Root "build\pyinstaller\$TargetArch"
    $distRoot = Join-Path $Root "dist"
    $commonArgs = @(
        "--noconfirm",
        "--clean",
        "--windowed",
        "--noupx",
        "--icon", $IconPath,
        "--add-data", "$Root\zadoo_vnc\templates;zadoo_vnc\templates",
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
        "pygrabber.dshow_graph",
        "soundcard",
        "sounddevice",
        "paramiko",
        "imagecodecs",
        "cv2"
    )
    if ($TargetArch -eq "x64") {
        $hiddenImports += @("av", "winpty", "dxcam", "bettercam")
        $winptyDir = (& $venvPython -c "import pathlib, winpty; print(pathlib.Path(winpty.__file__).parent)" | Select-Object -First 1).Trim()
        foreach ($winptyBinary in @("winpty-agent.exe", "OpenConsole.exe", "conpty.dll", "winpty.dll")) {
            $winptyPath = Join-Path $winptyDir $winptyBinary
            if (Test-Path -LiteralPath $winptyPath) {
                $commonArgs += @("--add-binary", "$winptyPath;winpty")
            } else {
                throw "Required pywinpty runtime file missing: $winptyPath"
            }
        }
    }
    foreach ($hiddenImport in $hiddenImports) {
        $commonArgs += @("--hidden-import", $hiddenImport)
    }
    $commonArgs += (Join-Path $Root "zadoo_vnc_single.py")

    $oneDirDist = Join-Path $distRoot "onedir\$TargetArch"
    & $venvPython -m PyInstaller @commonArgs --name "Zadoo" --onedir --distpath $oneDirDist --workpath (Join-Path $workRoot "onedir") --specpath (Join-Path $workRoot "spec-onedir")
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller one-folder build failed for $TargetArch" }
    $appExe = Join-Path $oneDirDist "Zadoo\Zadoo.exe"
    Invoke-Sign $appExe

    if (-not $SkipPortable) {
        $portableDist = Join-Path $distRoot "portable"
        & $venvPython -m PyInstaller @commonArgs --name "Zadoo-$TargetArch" --onefile --distpath $portableDist --workpath (Join-Path $workRoot "onefile") --specpath (Join-Path $workRoot "spec-onefile")
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller portable build failed for $TargetArch" }
        Invoke-Sign (Join-Path $portableDist "Zadoo-$TargetArch.exe")
    }

    if (-not $SkipInstaller) {
        $iscc = Resolve-Inno
        $installerOut = Join-Path $distRoot "installer"
        New-Item -ItemType Directory -Force -Path $installerOut | Out-Null
        $tempInstallerOut = Join-Path ([System.IO.Path]::GetTempPath()) ("zadoo-installer-$TargetArch-" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $tempInstallerOut | Out-Null
        try {
            & $iscc "/DSourceDir=$oneDirDist\Zadoo" "/DOutputDir=$tempInstallerOut" "/DArch=$TargetArch" "/DAppIcon=$IconPath" "/DAppVersion=$Version" (Join-Path $Root "installer\zadoo.iss")
            if ($LASTEXITCODE -ne 0) { throw "Inno Setup build failed for $TargetArch" }
            $builtInstaller = Join-Path $tempInstallerOut "Zadoo-$Version-$TargetArch-Setup.exe"
            if (-not (Test-Path -LiteralPath $builtInstaller)) { throw "Inno Setup did not create $builtInstaller" }
            Invoke-Sign $builtInstaller
            $finalInstaller = Join-Path $installerOut "Zadoo-$Version-$TargetArch-Setup.exe"
            $moved = $false
            for ($i = 1; $i -le 20; $i++) {
                try {
                    Move-Item -LiteralPath $builtInstaller -Destination $finalInstaller -Force
                    $moved = $true
                    break
                } catch {
                    if ($i -eq 20) { throw }
                    Start-Sleep -Milliseconds (250 * $i)
                }
            }
            if (-not $moved) { throw "Failed to move installer to $finalInstaller" }
        } finally {
            Remove-Item -LiteralPath $tempInstallerOut -Recurse -Force -ErrorAction SilentlyContinue
        }

        # The installer is fully self-contained. Remove the loose one-folder app and the
        # PyInstaller work directory so the ONLY artifact left is the installer itself
        # (no standalone Zadoo.exe). Pass -KeepOneDir to opt out.
        if (-not $KeepOneDir) {
            foreach ($leftover in @($oneDirDist, $workRoot)) {
                if ($leftover -and (Test-Path -LiteralPath $leftover)) {
                    Remove-Item -LiteralPath $leftover -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

Install-ToolsIfRequested
Set-Location $Root
Invoke-SourceChecks

$targets = if ($Arch -eq "All") { @("x64", "x86") } else { @($Arch) }
foreach ($target in $targets) {
    Build-Arch $target
}

Write-Step "Build complete"
Write-Host "Artifacts are under: $(Join-Path $Root 'dist')"
