param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if (-not (Test-Path ".\version.json")) {
    throw "version.json was not found."
}

$versionData = Get-Content ".\version.json" -Raw | ConvertFrom-Json
if ($Version -and $Version.Trim().Length -gt 0) {
    $versionData.version = $Version.Trim()
    $versionData | ConvertTo-Json | Set-Content ".\version.json" -Encoding UTF8
}

$appVersion = (Get-Content ".\version.json" -Raw | ConvertFrom-Json).version
$releaseDir = Join-Path $ProjectRoot "output\release"
$distDir = Join-Path $ProjectRoot "dist"
$buildDir = Join-Path $ProjectRoot "build"

if (Test-Path $distDir) { Remove-Item $distDir -Recurse -Force }
if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }
if (Test-Path $releaseDir) { Remove-Item $releaseDir -Recurse -Force }

python -m PyInstaller `
  --noconfirm `
  --name "MedicineUsageLoggingSystem" `
  --windowed `
  --clean `
  --add-data "version.json;." `
  GUI_test.py

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
Copy-Item ".\dist\MedicineUsageLoggingSystem\*" $releaseDir -Recurse -Force
Copy-Item ".\version.json" $releaseDir -Force

$zipPath = Join-Path $ProjectRoot "output\MedicineUsageLoggingSystem-v$appVersion.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

$zipSuccess = $false
for ($i = 0; $i -lt 3; $i++) {
    try {
        Compress-Archive -Path "$releaseDir\*" -DestinationPath $zipPath -CompressionLevel Optimal
        $zipSuccess = $true
        break
    }
    catch {
        Start-Sleep -Seconds 2
    }
}
if (-not $zipSuccess) {
    throw "Failed to create release zip package."
}

Write-Host "Build completed"
Write-Host "Version: $appVersion"
Write-Host "Release dir: $releaseDir"
Write-Host "Zip: $zipPath"
