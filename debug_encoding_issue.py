#!/usr/bin/env python3
"""Debug encoding issues in PyInstaller build process"""
import sys
import os
import subprocess
import pkg_resources
import tempfile

def debug_pkg_resources():
    """Debug pkg_resources encoding issues"""
    print("=== PKG_RESOURCES DEBUGGING ===")
    print(f"Python version: {sys.version}")
    print(f"Python executable: {sys.executable}")
    print(f"Current working directory: {os.getcwd()}")
    print(f"PYTHONPATH: {sys.path}")

    try:
        import pkg_resources
        print(f"pkg_resources location: {pkg_resources.__file__}")
        print(f"pkg_resources version: {pkg_resources.__version__}")

        # Check working_set
        print("\nInstalled packages (first 10):")
        for i, dist in enumerate(pkg_resources.working_set):
            if i >= 10:
                break
            print(f"  {dist.project_name}: {dist.version}")

    except Exception as e:
        print(f"ERROR accessing pkg_resources: {e}")
        return False

    return True

def debug_jaraco_text():
    """Debug jaraco.text module specifically"""
    print("\n=== JARACO.TEXT DEBUGGING ===")

    try:
        import jaraco.text
        print(f"jaraco.text location: {jaraco.text.__file__}")

        # Try to find the problematic file
        jaraco_init = jaraco.text.__file__
        print(f"Checking file: {jaraco_init}")

        # Read file with different encodings
        encodings_to_try = ['utf-8', 'windows-1252', 'latin-1', 'cp1252']

        for encoding in encodings_to_try:
            try:
                with open(jaraco_init, 'r', encoding=encoding) as f:
                    content = f.read(100)  # Read first 100 chars
                print(f"✓ {encoding}: Successfully read file")
                print(f"  First 50 chars: {repr(content[:50])}")
                break
            except UnicodeDecodeError as e:
                print(f"✗ {encoding}: {e}")

    except Exception as e:
        print(f"ERROR with jaraco.text: {e}")
        return False

    return True

def check_system_encoding():
    """Check system encoding settings"""
    print("\n=== SYSTEM ENCODING CHECK ===")

    print(f"Default encoding: {sys.getdefaultencoding()}")
    print(f"Filesystem encoding: {sys.getfilesystemencoding()}")
    print(f"Stdout encoding: {sys.stdout.encoding}")

    # Check environment variables
    encoding_vars = ['PYTHONIOENCODING', 'LANG', 'LC_ALL', 'LC_CTYPE']
    for var in encoding_vars:
        value = os.environ.get(var, 'Not set')
        print(f"{var}: {value}")

def test_pyinstaller_minimal():
    """Test minimal PyInstaller build"""
    print("\n=== MINIMAL PYINSTALLER TEST ===")

    # Create minimal test script
    test_script = """
import sys
print("Hello from minimal test")
print(f"Python version: {sys.version}")
"""

    with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as f:
        f.write(test_script)
        test_file = f.name

    try:
        # Try minimal PyInstaller build
        cmd = [
            sys.executable, '-m', 'PyInstaller',
            '--onefile',
            '--console',
            '--name=minimal_test',
            test_file
        ]

        print(f"Running: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

        print(f"Return code: {result.returncode}")
        if result.stdout:
            print(f"STDOUT:\n{result.stdout}")
        if result.stderr:
            print(f"STDERR:\n{result.stderr}")

        return result.returncode == 0

    except Exception as e:
        print(f"Minimal test failed: {e}")
        return False
    finally:
        try:
            os.unlink(test_file)
        except:
            pass

def main():
    print("CORPORATE ENCODING ISSUE DEBUGGER")
    print("=" * 50)

    debug_pkg_resources()
    debug_jaraco_text()
    check_system_encoding()

    print("\n" + "=" * 50)
    print("RECOMMENDATIONS:")

    # Check if we're in a corporate environment
    corporate_indicators = [
        'PROGRAMDATA' in os.environ and 'McAfee' in os.environ.get('PROGRAMDATA', ''),
        'PROGRAMFILES' in os.environ and any(antiv in os.environ['PROGRAMFILES'] for antiv in ['Symantec', 'Norton', 'McAfee']),
        os.path.exists('C:\\Program Files\\Cylance'),
        os.path.exists('C:\\ProgramData\\Sophos'),
    ]

    if any(corporate_indicators):
        print("🔒 CORPORATE ENVIRONMENT DETECTED")
        print("This encoding issue is likely caused by security software.")
        print("\nSOLUTIONS:")
        print("1. Use virtual environment with fresh package installs")
        print("2. Force UTF-8 encoding in environment variables")
        print("3. Use alternative PyInstaller settings")
        print("4. Build on personal machine and transfer")

    # Try minimal test
    print("\nRunning minimal PyInstaller test...")
    if test_pyinstaller_minimal():
        print("✓ Minimal PyInstaller works - issue is with specific packages")
    else:
        print("✗ Even minimal PyInstaller fails - system-level issue")

if __name__ == "__main__":
    main()
