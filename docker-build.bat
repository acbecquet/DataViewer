@echo off
echo Building Docker image for Testing GUI...
echo ========================================

REM Check if Docker is running
docker info >nul 2>&1
if errorlevel 1 (
    echo ERROR: Docker is not running. Please start Docker Desktop.
    pause
    exit /b 1
)

REM Build the Docker image
echo Building image...
docker build -t testing-gui:latest .

if errorlevel 1 (
    echo ERROR: Failed to build Docker image.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Docker image built successfully!
echo ========================================
pause
