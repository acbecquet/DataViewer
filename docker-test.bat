@echo off
echo Testing Docker GUI support...
echo ============================

REM Check if Docker is running
docker info >nul 2>&1
if errorlevel 1 (
    echo ERROR: Docker is not running. Please start Docker Desktop.
    pause
    exit /b 1
)

echo.
echo Testing simple GUI in Docker...
echo Make sure VcXsrv is running!
echo.

docker run -it --rm -e DISPLAY=host.docker.internal:0.0 testing-gui:latest python -c "import tkinter as tk; root = tk.Tk(); root.title('Docker Test'); root.geometry('300x200'); tk.Label(root, text='Success! GUI works!', font=('Arial', 16)).pack(expand=True); tk.Button(root, text='Close', command=root.quit).pack(); root.mainloop(); print('GUI test completed!')"

echo.
echo Test completed.
pause
