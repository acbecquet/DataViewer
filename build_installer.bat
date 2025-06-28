@echo off
echo ====================================================
echo    Professional Installer Builder
echo    Standardized Testing GUI v3.0.0
echo ====================================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo ERROR: PyInstaller not found. Installing...
    pip install pyinstaller
    if errorlevel 1 (
        echo FAILED to install PyInstaller
        pause
        exit /b 1
    )
)

REM Clean previous builds
echo [1/5] Cleaning previous builds...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "installer_output" rmdir /s /q "installer_output" 2>nul
echo ✓ Cleanup complete

REM Create missing directories
echo [2/5] Creating directories...
if not exist "resources" mkdir "resources"
if not exist "installer_output" mkdir "installer_output"

REM Create default icon if missing
if not exist "resources\icon.ico" (
    echo ⚠ Warning: No icon.ico found in resources\
    echo   Creating default icon...
    REM You can add icon creation here or just continue
)

REM Build executable with PyInstaller
echo [3/5] Building executable...
echo This may take 3-5 minutes...
pyinstaller testing_gui.spec --clean --noconfirm
if errorlevel 1 (
    echo ✗ PyInstaller build FAILED
    echo Check the output above for errors
    pause
    exit /b 1
)
echo ✓ Executable created successfully

REM Check if Inno Setup is available
echo [4/5] Checking for Inno Setup...
set "INNO_PATH=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%INNO_PATH%" (
    set "INNO_PATH=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)
if not exist "%INNO_PATH%" (
    echo.
    echo ⚠ Inno Setup not found!
    echo.
    echo Please install Inno Setup 6 from:
    echo https://jrsoftware.org/isdl.php
    echo.
    echo After installation, run this script again.
    echo.
    echo Manual steps:
    echo 1. Download and install Inno Setup
    echo 2. Run this script again
    echo 3. Your installer will be created automatically
    echo.
    pause
    exit /b 1
)

REM Create installer with Inno Setup
echo [5/5] Creating professional installer...
"%INNO_PATH%" "testing_gui_installer.iss"
if errorlevel 1 (
    echo ✗ Installer creation FAILED
    pause
    exit /b 1
)

echo.
echo ====================================================
echo    BUILD COMPLETE! 🎉
echo ====================================================
echo.
echo Your professional installer is ready:
dir "installer_output\*.exe" /b 2>nul
echo.
echo Location: installer_output\
echo.
echo What you can do now:
echo • Test the installer on a clean Windows machine
echo • Distribute the installer to users
echo • Users just double-click to install (no Python needed!)
echo.
echo The installer includes:
echo ✓ Your complete application
echo ✓ All required dependencies  
echo ✓ Start menu shortcut
echo ✓ Desktop icon (optional)
echo ✓ Professional uninstaller
echo ✓ File associations (.vap3 files)
echo.
pause