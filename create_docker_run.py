#!/usr/bin/env python3
"""Create docker-run.bat file safely without encryption issues."""

def create_docker_run_bat():
    """Create docker-run.bat with proper encoding."""
    
    content = '''@echo off
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
    -v "%cd%\\data:/app/user_data" ^
    -v "%cd%\\logs:/app/logs" ^
    -v "%cd%\\resources:/app/resources:ro" ^
    testing-gui:latest

echo.
echo Container stopped.
pause
'''
    
    # Write with ASCII encoding to avoid issues
    with open('docker-run.bat', 'w', encoding='ascii', newline='\r\n') as f:
        f.write(content)
    
    print("✓ Created docker-run.bat")

def create_docker_test_bat():
    """Create a simple test script."""
    
    content = '''@echo off
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
'''
    
    with open('docker-test.bat', 'w', encoding='ascii', newline='\r\n') as f:
        f.write(content)
    
    print("✓ Created docker-test.bat")

def create_docker_cleanup_bat():
    """Create cleanup script."""
    
    content = '''@echo off
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
'''
    
    with open('docker-cleanup.bat', 'w', encoding='ascii', newline='\r\n') as f:
        f.write(content)
    
    print("✓ Created docker-cleanup.bat")

def create_install_vcxsrv_guide():
    """Create a text guide for installing VcXsrv."""
    
    content = '''VcXsrv Installation Guide
========================

VcXsrv is required to display GUI applications from Docker containers on Windows.

Step 1: Download VcXsrv
-----------------------
Go to: https://sourceforge.net/projects/vcxsrv/
Click "Download" and save the installer.

Step 2: Install VcXsrv
----------------------
1. Run the downloaded installer
2. Follow the installation wizard
3. Accept default settings

Step 3: Configure XLaunch
-------------------------
1. Start "XLaunch" from the Start Menu
2. Display settings: Choose "Multiple windows"
3. Session type: Choose "Start no client"
4. Extra settings: 
   - CHECK "Disable access control"
   - CHECK "Native opengl" (optional)
5. Click "Next" then "Finish"

Step 4: Allow through Firewall
------------------------------
When Windows Firewall asks:
- Allow VcXsrv on both Private and Public networks

Step 5: Verify VcXsrv is Running
--------------------------------
- Look for VcXsrv icon in system tray (bottom-right corner)
- The icon should be visible when running

Step 6: Test Docker GUI
-----------------------
Run: docker-test.bat

This should open a small test window if everything is working.

Troubleshooting
---------------
If GUI doesn't appear:
1. Check VcXsrv is running (system tray icon)
2. Check Windows Firewall settings
3. Try restarting VcXsrv
4. Run docker-test.bat to isolate issues

For help, the application debug output will show any connection errors.
'''
    
    with open('VCXSRV_SETUP_GUIDE.txt', 'w', encoding='ascii') as f:
        f.write(content)
    
    print("✓ Created VCXSRV_SETUP_GUIDE.txt")

def main():
    print("Creating Docker batch files safely...")
    print("=" * 50)
    
    create_docker_run_bat()
    create_docker_test_bat()
    create_docker_cleanup_bat()
    create_install_vcxsrv_guide()
    
    print("\n" + "=" * 50)
    print("All Docker files created successfully!")
    print("=" * 50)
    print()
    print("Next steps:")
    print("1. Install VcXsrv (see VCXSRV_SETUP_GUIDE.txt)")
    print("2. Start XLaunch and configure it")
    print("3. Run: docker-test.bat (to test GUI)")
    print("4. Run: docker-run.bat (to start your app)")
    print()
    print("Files created:")
    print("  - docker-run.bat (main application)")
    print("  - docker-test.bat (simple GUI test)")
    print("  - docker-cleanup.bat (cleanup containers/images)")
    print("  - VCXSRV_SETUP_GUIDE.txt (installation guide)")

if __name__ == "__main__":
    main()