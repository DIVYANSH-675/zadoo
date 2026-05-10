param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$ToolsDir = Join-Path $Root "tools"
$FfmpegDir = Join-Path $ToolsDir "ffmpeg"
$MediaMtxDir = Join-Path $ToolsDir "mediamtx"
$TempDir = Join-Path $ToolsDir "_downloads"

New-Item -ItemType Directory -Force -Path $ToolsDir, $TempDir | Out-Null

function Get-GitHubAsset {
    param(
        [Parameter(Mandatory=$true)][string]$Repo,
        [Parameter(Mandatory=$true)][string]$AssetPattern
    )
    $headers = @{ "User-Agent" = "zadoo-turbo-setup" }
    $release = Invoke-RestMethod -Headers $headers -Uri "https://api.github.com/repos/$Repo/releases/latest"
    $asset = $release.assets | Where-Object { $_.name -match $AssetPattern } | Select-Object -First 1
    if (-not $asset) {
        throw "No asset matching '$AssetPattern' found in latest release for $Repo"
    }
    return $asset
}

function Download-File {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$OutFile
    )
    Write-Host "Downloading $Url"
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $OutFile
    $hash = Get-FileHash -Algorithm SHA256 -Path $OutFile
    Set-Content -Path "$OutFile.sha256" -Value $hash.Hash -Encoding ascii
    Write-Host "SHA256 $($hash.Hash)"
}

function Install-ZipTool {
    param(
        [Parameter(Mandatory=$true)][string]$ZipPath,
        [Parameter(Mandatory=$true)][string]$TargetDir,
        [Parameter(Mandatory=$true)][string]$ExeName
    )
    $extractDir = "$ZipPath.extract"
    if (Test-Path $extractDir) {
        Remove-Item -LiteralPath $extractDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $extractDir -Force

    $exe = Get-ChildItem -LiteralPath $extractDir -Recurse -Filter $ExeName | Select-Object -First 1
    if (-not $exe) {
        throw "$ExeName was not found inside $ZipPath"
    }

    if (Test-Path $TargetDir) {
        if ($Force) {
            Remove-Item -LiteralPath $TargetDir -Recurse -Force
        } else {
            Write-Host "$TargetDir already exists; use -Force to replace it."
        }
    }
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null

    $rootToCopy = $exe.Directory.FullName
    if ($ExeName -eq "ffmpeg.exe" -and (Split-Path $rootToCopy -Leaf) -eq "bin") {
        $rootToCopy = Split-Path $rootToCopy -Parent
    }
    Copy-Item -Path (Join-Path $rootToCopy "*") -Destination $TargetDir -Recurse -Force
    Remove-Item -LiteralPath $extractDir -Recurse -Force
}

function Assert-Executable {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Arg
    )
    if (-not (Test-Path $Path)) {
        throw "Missing executable: $Path"
    }
    $output = & $Path $Arg 2>&1 | Select-Object -First 5
    if ($LASTEXITCODE -ne 0 -and $Arg -eq "-version") {
        throw "$Path did not run successfully"
    }
    $output | ForEach-Object { Write-Host $_ }
}

$ffmpegExe = Join-Path $FfmpegDir "bin\ffmpeg.exe"
$mediaMtxExe = Join-Path $MediaMtxDir "mediamtx.exe"

if ($Force -or -not (Test-Path $ffmpegExe)) {
    $ffmpegZip = Join-Path $TempDir "ffmpeg-win64-gpl.zip"
    $ffmpegUrl = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    Download-File -Url $ffmpegUrl -OutFile $ffmpegZip
    Install-ZipTool -ZipPath $ffmpegZip -TargetDir $FfmpegDir -ExeName "ffmpeg.exe"
} else {
    Write-Host "FFmpeg already installed at $ffmpegExe"
}

if ($Force -or -not (Test-Path $mediaMtxExe)) {
    $asset = Get-GitHubAsset -Repo "bluenviron/mediamtx" -AssetPattern "windows_amd64\.zip$"
    $mediaZip = Join-Path $TempDir $asset.name
    Download-File -Url $asset.browser_download_url -OutFile $mediaZip
    Install-ZipTool -ZipPath $mediaZip -TargetDir $MediaMtxDir -ExeName "mediamtx.exe"
} else {
    Write-Host "MediaMTX already installed at $mediaMtxExe"
}

Write-Host ""
Write-Host "Validating FFmpeg..."
Assert-Executable -Path $ffmpegExe -Arg "-version"
Write-Host ""
Write-Host "Validating MediaMTX..."
Assert-Executable -Path $mediaMtxExe -Arg "--version"

Write-Host ""
Write-Host "Turbo streaming tools are ready."
Write-Host "Use these paths if you want explicit environment variables:"
Write-Host "`$env:ZADOO_FFMPEG_PATH='$ffmpegExe'"
Write-Host "`$env:ZADOO_MEDIAMTX_PATH='$mediaMtxExe'"
