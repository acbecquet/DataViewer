#!/usr/bin/env python3
"""Fix missing dependencies."""
import subprocess
import sys

def install_missing_deps():
    """Install missing dependencies."""
    missing_deps = [
        "requests>=2.25.0",
        "packaging>=20.0"  # Also needed if you used UpdateManager
    ]
    
    print("Installing missing dependencies...")
    for dep in missing_deps:
        print(f"Installing {dep}...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", dep], check=True)
            print(f"✓ {dep} installed successfully")
        except Exception as e:
            print(f"✗ Failed to install {dep}: {e}")

def test_imports():
    """Test that all imports work."""
    print("\nTesting imports...")
    
    try:
        import requests
        print("✓ requests import successful")
    except ImportError as e:
        print(f"✗ requests import failed: {e}")
    
    try:
        import main
        print("✓ main module import successful")
    except ImportError as e:
        print(f"✗ main module import failed: {e}")

def main():
    install_missing_deps()
    test_imports()
    
    print("\n" + "="*50)
    print("Now try running: testing-gui")
    print("="*50)

if __name__ == "__main__":
    main()