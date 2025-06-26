@echo off
echo Cleaning up Docker containers and images...
echo ==========================================

REM Stop any running containers
echo Stopping any running testing-gui containers...
docker stop testing-gui-container 2>nul

REM Remove stopped containers
echo Removing stopped containers...
docker container prune -f

REM Show current images
echo.
echo Current images:
docker images testing-gui

REM Ask if user wants to remove the image
echo.
set /p remove="Remove testing-gui image? (y/N): "
if /i "%remove%"=="y" (
    docker rmi testing-gui:latest
    echo Image removed.
) else (
    echo Image kept.
)

REM Clean up unused resources
echo.
echo Cleaning up unused Docker resources...
docker system prune -f

echo.
echo Cleanup completed!
pause
