@echo off
echo ====================================================
echo    Quick PyInstaller Build (No Inno Setup)
echo    Standardized Testing GUI v3.0.0
echo ====================================================
echo.

REM Clean previous builds
echo [1/3] Cleaning previous builds...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
echo ✓ Cleanup complete

REM Simple PyInstaller build without spec file complications
echo [2/3] Building executable (this may take 3-5 minutes)...
echo.

pyinstaller ^
    --name "TestingGUI" ^
    --windowed ^
    --onedir ^
    --clean ^
    --noconfirm ^
    --add-data "resources;resources" ^
    --hidden-import "sklearn.utils._cython_blas" ^
    --hidden-import "sklearn.neighbors.typedefs" ^
    --hidden-import "sklearn.tree._utils" ^
    --hidden-import "cv2" ^
    --hidden-import "PIL._tkinter_finder" ^
    --exclude-module "jupyter" ^
    --exclude-module "IPython" ^
    --exclude-module "notebook" ^
    --exclude-module "pytest" ^
    main.py

if errorlevel 1 (
    echo.
    echo ✗ Build FAILED
    echo.
    echo Common fixes:
    echo 1. Make sure all your Python files are in the current directory
    echo 2. Check that main.py exists and runs correctly: python main.py
    echo 3. Try installing missing packages: pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] Testing executable...
if exist "dist\TestingGUI\TestingGUI.exe" (
    echo ✓ Executable created successfully!
    echo.
    echo Location: dist\TestingGUI\TestingGUI.exe
    echo Size: 
    dir "dist\TestingGUI\TestingGUI.exe" | find "TestingGUI.exe"
    echo.
    echo ====================================================
    echo    BUILD COMPLETE! 🎉
    echo ====================================================
    echo.
    echo Your standalone application is ready:
    echo • Location: dist\TestingGUI\
    echo • Main file: TestingGUI.exe
    echo • No Python required to run!
    echo.
    echo To test:
    echo 1. Navigate to dist\TestingGUI\
    echo 2. Double-click TestingGUI.exe
    echo.
    echo To distribute:
    echo • Zip the entire dist\TestingGUI\ folder
    echo • Users extract and run TestingGUI.exe
    echo.
    set /p test="Test the executable now? (y/N): "
    if /i "%test%"=="y" (
        echo Starting TestingGUI.exe...
        start "" "dist\TestingGUI\TestingGUI.exe"
    )
    echo.
) else (
    echo ✗ Executable not found - build may have failed
    echo Check the output above for errors
)

pause