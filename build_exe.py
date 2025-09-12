#!/usr/bin/env python3
"""
build_exe.py
Creates a standalone executable using PyInstaller
Includes debugging output for troubleshooting
"""
import subprocess
import sys
import os
import shutil
from pathlib import Path

def cleanup_previous_builds():
    """Remove previous build artifacts"""
    print("DEBUG: Cleaning up previous builds...")

    # Remove previous builds
    dirs_to_remove = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"DEBUG: Removed directory: {dir_name}")

    # Remove spec files
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    for spec_file in spec_files:
        os.remove(spec_file)
        print(f"DEBUG: Removed spec file: {spec_file}")

def verify_main_file():
    """Verify main.py exists and is valid"""
    print("DEBUG: Verifying main.py...")

    if not os.path.exists('main.py'):
        print("ERROR: main.py not found!")
        return False

    try:
        with open('main.py', 'r') as f:
            content = f.read()
            if 'def main(' in content and 'TestingGUI' in content:
                print("DEBUG: main.py looks valid")
                return True
            else:
                print("ERROR: main.py missing required components")
                return False
    except Exception as e:
        print(f"ERROR: Could not read main.py: {e}")
        return False

def create_executable():
    """Create the executable using PyInstaller"""
    print("DEBUG: Creating executable with PyInstaller...")

    # PyInstaller command with comprehensive options
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',                    # Single executable file
        '--windowed',                   # No console window (GUI app)
        '--name=TestingGUI',           # Name of the executable
        '--icon=resources/icon.ico',    # Application icon
        '--add-data=resources;resources',  # Include resources folder
        '--hidden-import=matplotlib.backends.backend_tkagg',  # Ensure matplotlib backend
        '--hidden-import=PIL._tkinter_finder',               # PIL/Tkinter integration
        '--hidden-import=pkg_resources.py2_warn',            # Common missing import
        '--hidden-import=packaging.version',                 # Version checking
        '--hidden-import=packaging.specifiers',              # Package specifiers
        '--hidden-import=openpyxl.cell._writer',            # Excel writing
        '--hidden-import=tkintertable',                      # Table widget
        '--collect-all=matplotlib',     # Include all matplotlib
        '--collect-all=numpy',          # Include all numpy
        '--collect-all=pandas',         # Include all pandas
        '--exclude-module=pytest',      # Exclude test modules
        '--exclude-module=unittest',    # Exclude test modules
        '--exclude-module=doctest',     # Exclude test modules
        'main.py'
    ]

    print(f"DEBUG: Running command: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("DEBUG: PyInstaller completed successfully")
        print("DEBUG: PyInstaller stdout:", result.stdout[-500:])  # Last 500 chars
        return True
    except subprocess.CalledProcessError as e:
        print(f"ERROR: PyInstaller failed with return code {e.returncode}")
        print("ERROR: PyInstaller stderr:", e.stderr)
        print("ERROR: PyInstaller stdout:", e.stdout)
        return False

def verify_executable():
    """Verify the executable was created and works"""
    print("DEBUG: Verifying executable...")

    exe_path = 'dist/TestingGUI.exe'
    if not os.path.exists(exe_path):
        print(f"ERROR: Executable not found at {exe_path}")
        return False

    # Check file size (should be substantial for bundled app)
    size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"DEBUG: Executable size: {size_mb:.1f} MB")

    if size_mb < 50:  # Expect at least 50MB for bundled Python app
        print("WARNING: Executable seems too small, might be missing dependencies")

    print("DEBUG: Executable created successfully")
    return True

def copy_installer_files():
    """Copy additional files needed for installer"""
    print("DEBUG: Copying files for installer...")

    files_to_copy = [
        ('LICENSE.txt', 'dist/LICENSE.txt'),
        ('README.txt', 'dist/README.txt'),
    ]

    for src, dst in files_to_copy:
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"DEBUG: Copied {src} to {dst}")
        else:
            print(f"WARNING: {src} not found, skipping")

def main():
    """Main build process"""
    print("=" * 60)
    print("Building TestingGUI Executable")
    print("=" * 60)

    # Step 1: Cleanup
    cleanup_previous_builds()

    # Step 2: Verify main file
    if not verify_main_file():
        print("FAILED: Cannot proceed without valid main.py")
        return False

    # Step 3: Create executable
    if not create_executable():
        print("FAILED: Could not create executable")
        return False

    # Step 4: Verify executable
    if not verify_executable():
        print("FAILED: Executable verification failed")
        return False

    # Step 5: Copy additional files
    copy_installer_files()

    print("=" * 60)
    print("SUCCESS: Executable built successfully!")
    print("Next step: Run build_installer.bat to create the installer")
    print("=" * 60)

    return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)
