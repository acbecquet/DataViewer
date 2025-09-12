#!/usr/bin/env python3
"""Clean up failed installation attempts."""
import subprocess
import sys
import os
import shutil

def cleanup():
    """Clean up build artifacts and failed installations."""
    print("Cleaning up...")

    # Remove build artifacts
    dirs_to_clean = ['build', 'dist', 'standardized_testing_gui.egg-info']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Removed: {dir_name}")

    # Try to uninstall any partial installation
    try:
        subprocess.run([sys.executable, "-m", "pip", "uninstall", "standardized-testing-gui", "-y"],
                      capture_output=True)
        print("Removed any existing installation")
    except:
        pass

    print("Cleanup completed!")

if __name__ == "__main__":
    cleanup()
