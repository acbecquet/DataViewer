@echo off
echo ================================================
echo Building TestingGUI Professional Installer
echo ================================================

REM Check if Inno Setup is installed
if not exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo ERROR: Inno Setup 6 not found!
    echo Please install from: https://jrsoftware.org/isdl.php
    pause
    exit /b 1
)

REM Step 1: Build the executable
echo.
echo Step 1: Building executable...
python build_exe.py
if errorlevel 1 (
    echo ERROR: Failed to build executable
    pause
    exit /b 1
)

REM Step 2: Create the installer
echo.
echo Step 2: Creating installer with Inno Setup...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_script.iss
if errorlevel 1 (
    echo ERROR: Failed to create installer
    pause
    exit /b 1
)

echo.
echo ================================================
echo SUCCESS: Professional installer created!
echo Location: installer_output\TestingGUI_Setup_v3.0.0.exe
echo ================================================
echo.
echo Testing installer...
dir installer_output\*.exe

echo.
echo Build complete! You can now distribute the installer.
pause