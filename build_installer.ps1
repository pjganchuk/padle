<# 
PADLE Build Script
==================
Builds PADLE into a Windows installer.

Prerequisites:
1. Python 3.10+ with venv activated
2. PyInstaller installed: pip install pyinstaller
3. Inno Setup 6.x installed (https://jrsoftware.org/isinfo.php)
4. Tesseract OCR portable (will be downloaded if missing)

Usage:
  .\build_installer.ps1
  .\build_installer.ps1 -SkipPyInstaller    # Skip PyInstaller, just build installer
  .\build_installer.ps1 -Clean              # Clean build directories first
#>

param(
    [switch]$SkipPyInstaller,
    [switch]$Clean,
    [switch]$Help
)

# Configuration
$AppName = "PADLE"
$AppVersion = "1.1.0"
$TesseractVersion = "5.3.3"
$TesseractUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v${TesseractVersion}/tesseract-ocr-w64-setup-${TesseractVersion}.20231005.exe"

# Colors for output
function Write-Step { param($msg) Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Write-Success { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warning { param($msg) Write-Host "    [!] $msg" -ForegroundColor Yellow }
function Write-Error { param($msg) Write-Host "    [ERROR] $msg" -ForegroundColor Red }

# Show help
if ($Help) {
    Write-Host @"

PADLE Build Script
==================

This script builds PADLE into a Windows installer.

Prerequisites:
  1. Python virtual environment activated with all dependencies
  2. PyInstaller: pip install pyinstaller  
  3. Inno Setup 6.x: https://jrsoftware.org/isinfo.php

Options:
  -SkipPyInstaller  Skip the PyInstaller step (use existing dist folder)
  -Clean            Remove build directories before starting
  -Help             Show this help message

Steps performed:
  1. Check prerequisites (Python, PyInstaller, Inno Setup)
  2. Download Tesseract OCR portable (if not present)
  3. Run PyInstaller to create executable
  4. Run Inno Setup to create installer

Output:
  installer_output\PADLE_Setup_$AppVersion.exe

"@
    exit 0
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PADLE Build Script v$AppVersion" -ForegroundColor Cyan  
Write-Host "========================================" -ForegroundColor Cyan

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = Get-Location }
Set-Location $ScriptDir

# Clean if requested
if ($Clean) {
    Write-Step "Cleaning build directories..."
    
    if (Test-Path "build") { 
        Remove-Item -Recurse -Force "build"
        Write-Success "Removed build/"
    }
    if (Test-Path "dist") {
        Remove-Item -Recurse -Force "dist" 
        Write-Success "Removed dist/"
    }
    if (Test-Path "installer_output") {
        Remove-Item -Recurse -Force "installer_output"
        Write-Success "Removed installer_output/"
    }
}

# =============================================================================
# Step 1: Check Prerequisites
# =============================================================================
Write-Step "Checking prerequisites..."

# Check Python
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "Python not found. Please activate your virtual environment."
    exit 1
}
Write-Success "Python: $pythonVersion"

# Check if in venv
if (-not $env:VIRTUAL_ENV) {
    Write-Warning "Virtual environment not detected. Make sure dependencies are installed."
}

# Check PyInstaller
$pyinstallerVersion = pyinstaller --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller not found. Install with: pip install pyinstaller"
    exit 1
}
Write-Success "PyInstaller: $pyinstallerVersion"

# Check Inno Setup
$InnoSetupPath = $null
$InnoPaths = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)

foreach ($path in $InnoPaths) {
    if (Test-Path $path) {
        $InnoSetupPath = $path
        break
    }
}

if (-not $InnoSetupPath) {
    Write-Error "Inno Setup not found. Download from: https://jrsoftware.org/isinfo.php"
    exit 1
}
Write-Success "Inno Setup: $InnoSetupPath"

# Check for icon files
if (-not (Test-Path "icon.ico")) {
    Write-Warning "icon.ico not found - installer will use default icon"
    Write-Host "         To create icon.ico from icon.png, use ImageMagick:" -ForegroundColor Gray
    Write-Host "         magick icon.png -define icon:auto-resize=256,128,64,48,32,16 icon.ico" -ForegroundColor Gray
}

if (-not (Test-Path "icon.png")) {
    Write-Warning "icon.png not found"
}

# =============================================================================
# Step 2: Download/Setup Tesseract OCR
# =============================================================================
Write-Step "Setting up Tesseract OCR..."

$TesseractDir = Join-Path $ScriptDir "tesseract"

if (Test-Path (Join-Path $TesseractDir "tesseract.exe")) {
    Write-Success "Tesseract already present"
} else {
    Write-Host "    Downloading Tesseract OCR portable..." -ForegroundColor Gray
    
    # Create temp directory
    $TempDir = Join-Path $env:TEMP "padle_build"
    New-Item -ItemType Directory -Force -Path $TempDir | Out-Null
    
    $InstallerPath = Join-Path $TempDir "tesseract_installer.exe"
    
    try {
        # Download installer
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $TesseractUrl -OutFile $InstallerPath -UseBasicParsing
        
        Write-Host "    Extracting Tesseract (this runs the installer silently)..." -ForegroundColor Gray
        
        # Run installer silently to extract to our directory
        # Note: Tesseract installer doesn't support true portable extraction
        # We'll install to a temp location then copy
        $TempInstall = Join-Path $TempDir "tesseract_install"
        Start-Process -FilePath $InstallerPath -ArgumentList "/S", "/D=$TempInstall" -Wait
        
        # Copy to our tesseract folder
        if (Test-Path $TempInstall) {
            New-Item -ItemType Directory -Force -Path $TesseractDir | Out-Null
            Copy-Item -Path "$TempInstall\*" -Destination $TesseractDir -Recurse -Force
            Write-Success "Tesseract extracted to tesseract/"
        } else {
            Write-Error "Tesseract extraction failed"
            Write-Host "         Please download manually from:" -ForegroundColor Gray
            Write-Host "         https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Gray
        }
        
    } catch {
        Write-Error "Failed to download Tesseract: $_"
        Write-Host "         Please download manually from:" -ForegroundColor Gray
        Write-Host "         https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Gray
    }
    
    # Cleanup
    Remove-Item -Recurse -Force $TempDir -ErrorAction SilentlyContinue
}

# =============================================================================
# Step 3: Run PyInstaller
# =============================================================================
if (-not $SkipPyInstaller) {
    Write-Step "Running PyInstaller..."
    
    if (-not (Test-Path "padle.spec")) {
        Write-Error "padle.spec not found"
        exit 1
    }
    
    # Run PyInstaller
    pyinstaller padle.spec --noconfirm
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "PyInstaller failed"
        exit 1
    }
    
    # Verify output
    if (Test-Path "dist\PADLE\PADLE.exe") {
        Write-Success "PyInstaller completed: dist\PADLE\"
        
        # Show size
        $size = (Get-ChildItem -Path "dist\PADLE" -Recurse | Measure-Object -Property Length -Sum).Sum
        $sizeMB = [math]::Round($size / 1MB, 1)
        Write-Host "         Distribution size: ${sizeMB} MB" -ForegroundColor Gray
    } else {
        Write-Error "PADLE.exe not found in dist folder"
        exit 1
    }
} else {
    Write-Step "Skipping PyInstaller (using existing dist folder)..."
    
    if (-not (Test-Path "dist\PADLE\PADLE.exe")) {
        Write-Error "dist\PADLE\PADLE.exe not found. Run without -SkipPyInstaller first."
        exit 1
    }
    Write-Success "Using existing dist\PADLE\"
}

# =============================================================================
# Step 4: Run Inno Setup
# =============================================================================
Write-Step "Running Inno Setup..."

if (-not (Test-Path "padle_installer.iss")) {
    Write-Error "padle_installer.iss not found"
    exit 1
}

# Create output directory
New-Item -ItemType Directory -Force -Path "installer_output" | Out-Null

# Run Inno Setup
& $InnoSetupPath "padle_installer.iss"

if ($LASTEXITCODE -ne 0) {
    Write-Error "Inno Setup failed"
    exit 1
}

# Find output file
$InstallerFile = Get-ChildItem -Path "installer_output" -Filter "PADLE_Setup_*.exe" | Select-Object -First 1

if ($InstallerFile) {
    Write-Success "Installer created: $($InstallerFile.FullName)"
    
    $sizeMB = [math]::Round($InstallerFile.Length / 1MB, 1)
    Write-Host "         Installer size: ${sizeMB} MB" -ForegroundColor Gray
} else {
    Write-Error "Installer file not found in installer_output/"
    exit 1
}

# =============================================================================
# Done!
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Build Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Installer: installer_output\$($InstallerFile.Name)" -ForegroundColor White
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Yellow
Write-Host "  1. Test the installer on a clean Windows machine" -ForegroundColor Gray
Write-Host "  2. Verify PADLE launches and can generate descriptions" -ForegroundColor Gray
Write-Host "  3. Test with and without VLC/FFmpeg installed" -ForegroundColor Gray
Write-Host ""
