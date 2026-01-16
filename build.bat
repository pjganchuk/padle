@echo off
REM PADLE Build Script
REM Builds PADLE into a distributable Windows application
REM
REM Prerequisites:
REM   - Python virtual environment activated
REM   - PyInstaller installed: pip install pyinstaller
REM
REM Output: dist\PADLE\PADLE.exe

echo.
echo ========================================
echo   PADLE Build Script
echo ========================================
echo.

REM Check if virtual environment is activated
if "%VIRTUAL_ENV%"=="" (
    echo ERROR: Virtual environment not activated!
    echo.
    echo Please run:
    echo   .\venv\Scripts\activate
    echo.
    echo Then run this script again.
    pause
    exit /b 1
)

REM Check if PyInstaller is installed
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

REM Check for icon files
if not exist "icon.png" (
    echo WARNING: icon.png not found
)
if not exist "icon.ico" (
    echo WARNING: icon.ico not found
    echo.
    echo To create icon.ico from icon.png, you can use:
    echo   - Online converter: https://convertio.co/png-ico/
    echo   - ImageMagick: magick icon.png -define icon:auto-resize=256,128,64,48,32,16 icon.ico
    echo.
)

REM Clean previous build
echo Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist\PADLE" rmdir /s /q dist\PADLE

REM Run PyInstaller
echo.
echo Building PADLE...
echo.
pyinstaller padle.spec --noconfirm

if errorlevel 1 (
    echo.
    echo BUILD FAILED!
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Build Complete!
echo ========================================
echo.
echo Output: dist\PADLE\PADLE.exe
echo.
echo To test the build:
echo   1. Start Moondream in a separate terminal
echo   2. Run dist\PADLE\PADLE.exe
echo.
echo To distribute:
echo   - Copy the entire dist\PADLE folder
echo   - Users will also need VLC, Tesseract, and FFmpeg installed
echo.
pause
