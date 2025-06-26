#!/usr/bin/env python3
"""Fix missing imports quickly."""
import subprocess
import sys

def add_missing_imports():
    """Add missing imports to common files."""
    
    # Fix utils.py
    print("Fixing utils.py...")
    with open("utils.py", "r") as f:
        content = f.read()
    
    if "import traceback" not in content:
        # Find the import section and add traceback
        lines = content.split('\n')
        import_index = -1
        
        for i, line in enumerate(lines):
            if line.startswith('import ') and 'traceback' not in line:
                import_index = i
            elif line.startswith('from ') or (line.strip() == '' and import_index > -1):
                break
        
        if import_index > -1:
            lines.insert(import_index + 1, "import traceback")
            
            with open("utils.py", "w") as f:
                f.write('\n'.join(lines))
            print("✓ Added traceback import to utils.py")
        else:
            print("⚠ Could not automatically add traceback import")
    else:
        print("✓ traceback already imported in utils.py")
    
    # Check for other common missing imports
    missing_deps = []
    
    try:
        import cv2
        print("✓ opencv-python available")
    except ImportError:
        missing_deps.append("opencv-python")
        print("⚠ opencv-python missing")
    
    try:
        import requests
        print("✓ requests available")
    except ImportError:
        missing_deps.append("requests")
        print("⚠ requests missing")
    
    # Install missing dependencies
    if missing_deps:
        print(f"\nInstalling missing dependencies: {missing_deps}")
        for dep in missing_deps:
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", dep], check=True)
                print(f"✓ Installed {dep}")
            except Exception as e:
                print(f"✗ Failed to install {dep}: {e}")

def main():
    print("Fixing missing imports...")
    add_missing_imports()
    print("\nTry running 'testing-gui' again!")

if __name__ == "__main__":
    main()