@echo off
echo Starting Testing GUI in Docker...
echo =================================

REM Check if Docker is running
docker info >nul 2>&1
if errorlevel 1 (
    echo ERROR: Docker is not running. Please start Docker Desktop.
    pause
    exit /b 1
)

REM Create data directories if they don't exist
if not exist "data" mkdir data
if not exist "logs" mkdir logs

echo.
echo NOTE: Make sure VcXsrv is running for GUI display!
echo You should see VcXsrv icon in your system tray.
echo.
echo If you don't have VcXsrv:
echo 1. Download from: https://sourceforge.net/projects/vcxsrv/
echo 2. Install and run XLaunch
echo 3. Choose "Multiple windows" and "Disable access control"
echo.

REM Run the container with GUI support
echo Starting container...
docker run -it --rm ^
    --name testing-gui-container ^
    -e DISPLAY=host.docker.internal:0.0 ^
    -v "%cd%\data:/app/user_data" ^
    -v "%cd%\logs:/app/logs" ^
    -v "%cd%\resources:/app/resources:ro" ^
    testing-gui:latest

echo.
echo Container stopped.
pause
