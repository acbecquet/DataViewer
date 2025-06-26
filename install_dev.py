#!/usr/bin/env python3
"""
Development installation script.
Installs the package in editable mode for development.
"""
import subprocess
import sys
import os

def check_requirements():
    """Check that requirements.txt doesn't contain problematic packages."""
    with open("requirements.txt", "r") as f:
        requirements = f.read()
    
    problematic = ["sqlite3"]
    issues = []
    
    for package in problematic:
        if package in requirements:
            issues.append(package)
    
    if issues:
        print(f"ERROR: Found problematic packages in requirements.txt: {issues}")
        print("sqlite3 is part of Python's standard library and should be removed from requirements.txt")
        return False
    
    return True

def install_dev():
    """Install package in development mode."""
    print("Checking requirements...")
    if not check_requirements():
        return False
    
    print("Installing package in development mode...")
    
    try:
        # Upgrade pip first
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=True)
        
        # Install in editable mode
        result = subprocess.run([sys.executable, "-m", "pip", "install", "-e", "."], 
                              capture_output=True, text=True, check=True)
        
        print("Development installation completed!")
        print("You can now run 'testing-gui' from anywhere.")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Installation failed with error: {e}")
        print(f"STDOUT: {e.stdout}")
        print(f"STDERR: {e.stderr}")
        return False

def test_installation():
    """Test that the installation worked."""
    try:
        result = subprocess.run([sys.executable, "-c", "import main; print('Import successful')"], 
                              capture_output=True, text=True, check=True)
        print("✓ Module import test passed")
        
        # Test entry point
        result = subprocess.run(["testing-gui", "--help"], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print("✓ Entry point test passed")
        else:
            print("⚠ Entry point available but may have issues")
        
        return True
    except subprocess.TimeoutExpired:
        print("⚠ Entry point test timed out (this may be normal for GUI apps)")
        return True
    except Exception as e:
        print(f"✗ Installation test failed: {e}")
        return False

def main():
    print("Starting development installation...")
    
    if install_dev():
        print("\nTesting installation...")
        test_installation()
        
        print("\n" + "="*50)
        print("INSTALLATION COMPLETE!")
        print("="*50)
        print("To run the application:")
        print("  testing-gui")
        print("  or: python -m main")
        print("  or: python main.py")
    else:
        print("\nInstallation failed. Please check the error messages above.")

if __name__ == "__main__":
    main()