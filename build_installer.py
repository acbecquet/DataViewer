# build_installer.py
import os
import sys
import subprocess
from pathlib import Path

def build_for_platform():
    """Build platform-specific packages."""
    platform = sys.platform

    if platform == "win32":
        # Windows: Create MSI installer
        subprocess.run([
            "python", "setup.py", "bdist_msi"
        ])
    elif platform == "darwin":
        # macOS: Create app bundle
        subprocess.run([
            "python", "setup.py", "bdist_dmg"
        ])
    else:
        # Linux: Create AppImage or deb package
        subprocess.run([
            "python", "setup.py", "bdist_rpm"
        ])

if __name__ == "__main__":
    build_for_platform()
