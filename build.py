#!/usr/bin/env python3
"""
Build script for the Testing GUI application.
Run this to create distribution packages.
"""
import subprocess
import sys
import os
import shutil

def clean_build():
    """Clean previous build artifacts."""
    dirs_to_clean = ['build', 'dist', '*.egg-info']
    for dir_pattern in dirs_to_clean:
        if '*' in dir_pattern:
            import glob
            for path in glob.glob(dir_pattern):
                if os.path.exists(path):
                    shutil.rmtree(path)
                    print(f"Cleaned: {path}")
        else:
            if os.path.exists(dir_pattern):
                shutil.rmtree(dir_pattern)
                print(f"Cleaned: {dir_pattern}")

def build_package():
    """Build the package."""
    print("Building package...")
    
    # Build source distribution
    subprocess.run([sys.executable, "setup.py", "sdist"], check=True)
    
    # Build wheel distribution
    subprocess.run([sys.executable, "setup.py", "bdist_wheel"], check=True)
    
    print("Build completed! Check the 'dist' directory for packages.")

def main():
    print("Starting build process...")
    clean_build()
    build_package()
    
    print("\nTo install locally for testing:")
    print("pip install -e .")
    print("\nTo install from wheel:")
    print("pip install dist/standardized_testing_gui-3.0.0-py3-none-any.whl")

if __name__ == "__main__":
    main()