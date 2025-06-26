#!/usr/bin/env python3
"""
Development installation script.
Installs the package in editable mode for development.
"""
import subprocess
import sys

def install_dev():
    """Install package in development mode."""
    print("Installing package in development mode...")
    
    # Install in editable mode
    subprocess.run([sys.executable, "-m", "pip", "install", "-e", "."], check=True)
    
    print("Development installation completed!")
    print("You can now run 'testing-gui' from anywhere.")

if __name__ == "__main__":
    install_dev()