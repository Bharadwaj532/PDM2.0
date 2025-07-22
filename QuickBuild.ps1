# Quick Build Script for Package Development Manager
# Creates executable using ps2exe

Write-Host "Package Development Manager - Quick Build" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

# Check if ps2exe is available
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing ps2exe module..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}

# Check if main file exists
if (-not (Test-Path "PackageDevManager_GUI.ps1")) {
    Write-Error "PackageDevManager_GUI.ps1 not found!"
    exit 1
}

Write-Host "Building PackageDevManager.exe..." -ForegroundColor Cyan

# Look for icon
$icon = $null
if (Test-Path "PackageDevManager.ico") { $icon = "PackageDevManager.ico" }
elseif (Test-Path "icon.ico") { $icon = "icon.ico" }

# Build parameters
$params = @{
    inputFile = "PackageDevManager_GUI.ps1"
    outputFile = "PackageDevManager.exe"
    title = "Package Development Manager"
    description = "Optum Package Development Manager v1.5"
    company = "Optum"
    product = "Package Development Manager"
    copyright = "© 2024 Optum"
    version = "1.5.0.0"
    requireAdmin = $true
}

if ($icon) {
    $params.iconFile = $icon
    Write-Host "Using icon: $icon" -ForegroundColor Green
}

# Convert to EXE
Invoke-ps2exe @params

# Check result
if (Test-Path "PackageDevManager.exe") {
    $size = [math]::Round((Get-Item "PackageDevManager.exe").Length / 1MB, 2)
    Write-Host ""
    Write-Host "SUCCESS! PackageDevManager.exe created ($size MB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Usage: PackageDevManager.exe" -ForegroundColor Cyan
    Write-Host ""
} else {
    Write-Error "Build failed!"
}
