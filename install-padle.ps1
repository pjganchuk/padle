<# 
PADLE Installer for Windows
Panopto Audio Descriptions List Editor

This script will:
1. Check for Python installation
2. Download and install Tesseract OCR
3. Check for VLC installation
4. Create a virtual environment
5. Install Python dependencies
6. Provide instructions for Moondream
#>

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PADLE Installer" -ForegroundColor Cyan
Write-Host "  Panopto Audio Descriptions List Editor" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get the directory where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($scriptDir)) {
    $scriptDir = Get-Location
}

Write-Host "Installing to: $scriptDir" -ForegroundColor Gray
Write-Host ""

# ============================================
# STEP 1: Check for Python
# ============================================
Write-Host "[Step 1/6] Checking for Python..." -ForegroundColor Yellow

$pythonCmd = $null
$pythonVersion = $null

# Try different Python commands
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $version = & $cmd --version 2>&1
        if ($version -match "Python (\d+)\.(\d+)") {
            $major = [int]$Matches[1]
            $minor = [int]$Matches[2]
            if ($major -ge 3 -and $minor -ge 8) {
                $pythonCmd = $cmd
                $pythonVersion = $version
                break
            }
        }
    } catch {
        # Command not found, try next
    }
}

if ($pythonCmd) {
    Write-Host "  Found: $pythonVersion" -ForegroundColor Green
} else {
    Write-Host "  Python 3.8+ not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Please install Python 3.10 or later from:" -ForegroundColor White
    Write-Host "  https://www.python.org/downloads/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  IMPORTANT: Check 'Add Python to PATH' during installation!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  After installing Python, run this script again." -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# ============================================
# STEP 2: Check/Install Tesseract OCR
# ============================================
Write-Host ""
Write-Host "[Step 2/6] Checking for Tesseract OCR..." -ForegroundColor Yellow

$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$tesseractInstalled = Test-Path $tesseractPath

if ($tesseractInstalled) {
    Write-Host "  Found: Tesseract OCR" -ForegroundColor Green
} else {
    Write-Host "  Tesseract OCR not found. Downloading installer..." -ForegroundColor White
    
    # Download Tesseract installer from UB-Mannheim
    $tesseractUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    $installerPath = "$env:TEMP\tesseract-installer.exe"
    
    try {
        Write-Host "  Downloading from GitHub..." -ForegroundColor Gray
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $tesseractUrl -OutFile $installerPath -UseBasicParsing
        
        Write-Host "  Running installer..." -ForegroundColor Gray
        Write-Host ""
        Write-Host "  >>> A Tesseract installer window will open. <<<" -ForegroundColor Yellow
        Write-Host "  >>> Use the DEFAULT installation path! <<<" -ForegroundColor Yellow
        Write-Host ""
        
        # Run installer (not silent, so user can see progress)
        Start-Process -FilePath $installerPath -Wait
        
        # Verify installation
        if (Test-Path $tesseractPath) {
            Write-Host "  Tesseract installed successfully!" -ForegroundColor Green
        } else {
            Write-Host "  Warning: Tesseract may not have installed correctly." -ForegroundColor Yellow
            Write-Host "  Please install manually from:" -ForegroundColor White
            Write-Host "  https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Cyan
        }
        
        # Clean up
        Remove-Item $installerPath -ErrorAction SilentlyContinue
        
    } catch {
        Write-Host "  Failed to download Tesseract: $_" -ForegroundColor Red
        Write-Host "  Please install manually from:" -ForegroundColor White
        Write-Host "  https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Cyan
    }
}

# ============================================
# STEP 3: Check for VLC
# ============================================
Write-Host ""
Write-Host "[Step 3/6] Checking for VLC Media Player..." -ForegroundColor Yellow

$vlcPaths = @(
    "C:\Program Files\VideoLAN\VLC\vlc.exe",
    "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"
)

$vlcInstalled = $false
foreach ($path in $vlcPaths) {
    if (Test-Path $path) {
        $vlcInstalled = $true
        break
    }
}

if ($vlcInstalled) {
    Write-Host "  Found: VLC Media Player" -ForegroundColor Green
} else {
    Write-Host "  VLC not found." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  VLC is required for audio playback." -ForegroundColor White
    Write-Host "  Please download and install from:" -ForegroundColor White
    Write-Host "  https://www.videolan.org/vlc/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  After installing VLC, press Enter to continue..." -ForegroundColor Yellow
    Read-Host
}

# ============================================
# STEP 4: Create Virtual Environment
# ============================================
Write-Host ""
Write-Host "[Step 4/6] Setting up Python virtual environment..." -ForegroundColor Yellow

$venvPath = Join-Path $scriptDir "venv"

if (Test-Path $venvPath) {
    Write-Host "  Virtual environment already exists." -ForegroundColor Green
} else {
    Write-Host "  Creating virtual environment..." -ForegroundColor Gray
    & $pythonCmd -m venv $venvPath
    
    if (Test-Path $venvPath) {
        Write-Host "  Virtual environment created!" -ForegroundColor Green
    } else {
        Write-Host "  Failed to create virtual environment!" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# ============================================
# STEP 5: Install Python Dependencies
# ============================================
Write-Host ""
Write-Host "[Step 5/6] Installing Python dependencies..." -ForegroundColor Yellow

$pipPath = Join-Path $venvPath "Scripts\pip.exe"
$requirementsPath = Join-Path $scriptDir "requirements.txt"

if (Test-Path $requirementsPath) {
    Write-Host "  Installing from requirements.txt..." -ForegroundColor Gray
    & $pipPath install -r $requirementsPath
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Dependencies installed!" -ForegroundColor Green
    } else {
        Write-Host "  Some dependencies may have failed to install." -ForegroundColor Yellow
    }
} else {
    Write-Host "  requirements.txt not found, installing manually..." -ForegroundColor Gray
    & $pipPath install opencv-python numpy Pillow moondream pytesseract screeninfo python-vlc
}

# Install moondream-station
Write-Host "  Installing moondream-station..." -ForegroundColor Gray
& $pipPath install moondream-station

# ============================================
# STEP 6: Final Instructions
# ============================================
Write-Host ""
Write-Host "[Step 6/6] Final Setup - Moondream AI Server" -ForegroundColor Yellow
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "To run PADLE, you need to start the Moondream AI server first." -ForegroundColor White
Write-Host ""
Write-Host "EVERY TIME you want to use PADLE:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Open a Command Prompt or PowerShell" -ForegroundColor White
Write-Host ""
Write-Host "  2. Navigate to this folder:" -ForegroundColor White
Write-Host "     cd `"$scriptDir`"" -ForegroundColor Cyan
Write-Host ""
Write-Host "  3. Activate the virtual environment:" -ForegroundColor White
Write-Host "     .\venv\Scripts\activate" -ForegroundColor Cyan
Write-Host ""
Write-Host "  4. Start Moondream:" -ForegroundColor White
Write-Host "     moondream-station" -ForegroundColor Cyan
Write-Host ""
Write-Host "  5. In the Moondream terminal, switch to moondream-2:" -ForegroundColor White
Write-Host "     models switch moondream-2" -ForegroundColor Cyan
Write-Host ""
Write-Host "  6. Double-click 'run-padle.bat' or run:" -ForegroundColor White
Write-Host "     python main.py" -ForegroundColor Cyan
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Create a batch file to run the app easily
$batchPath = Join-Path $scriptDir "run-padle.bat"
$batchContent = @"
@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python main.py
"@
Set-Content -Path $batchPath -Value $batchContent
Write-Host "Created 'run-padle.bat' for easy launching." -ForegroundColor Green
Write-Host ""

Read-Host "Press Enter to exit"
