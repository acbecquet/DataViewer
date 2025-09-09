#!/usr/bin/env python3
"""Progressive build test to find problematic packages"""

import sys
import os
import subprocess
import time

def test_build_with_packages(packages_to_include, test_name):
    """Test PyInstaller build with specific packages"""
    print(f"\n{'='*60}")
    print(f"TESTING BUILD: {test_name}")
    print(f"Packages: {packages_to_include}")
    print('='*60)
    
    # Create test script
    test_script = f'''
import sys
print("Testing with packages: {packages_to_include}")

# Try importing the packages
try:
'''
    
    for pkg in packages_to_include:
        if pkg == 'matplotlib':
            test_script += '''    import matplotlib
    matplotlib.use('TkAgg')
    print("✓ matplotlib imported")
'''
        elif pkg == 'pandas':
            test_script += '''    import pandas as pd
    print("✓ pandas imported")
'''
        elif pkg == 'numpy':
            test_script += '''    import numpy as np
    print("✓ numpy imported")
'''
        elif pkg == 'PIL':
            test_script += '''    from PIL import Image
    print("✓ PIL imported")
'''
        elif pkg == 'pptx':
            test_script += '''    from pptx import Presentation
    print("✓ pptx imported")
'''
        elif pkg == 'openpyxl':
            test_script += '''    import openpyxl
    print("✓ openpyxl imported")
'''
    
    test_script += '''
    print("All imports successful!")
except Exception as e:
    print(f"Import error: {e}")
    sys.exit(1)

print("Test completed successfully")
'''
    
    # Write test file
    test_file = f'test_{test_name.lower().replace(" ", "_")}.py'
    with open(test_file, 'w') as f:
        f.write(test_script)
    
    # Build PyInstaller command
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--console',
        f'--name=test_{test_name.lower().replace(" ", "_")}',
        '--clean',
        '--noconfirm'
    ]
    
    # Add collect-all for each package
    for pkg in packages_to_include:
        cmd.append(f'--collect-all={pkg}')
    
    # Add excludes for known problematic modules
    excludes = ['pkg_resources', 'setuptools', 'distutils', 'tkinter.test', 'test']
    for exclude in excludes:
        cmd.append(f'--exclude-module={exclude}')
    
    cmd.append(test_file)
    
    print(f"Command: {' '.join(cmd)}")
    
    try:
        start_time = time.time()
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        end_time = time.time()
        
        print(f"Build time: {end_time - start_time:.1f} seconds")
        print(f"Return code: {result.returncode}")
        
        if result.returncode == 0:
            print(f"✅ SUCCESS: {test_name}")
            
            # Check if exe exists and try to run it
            exe_path = f'dist/test_{test_name.lower().replace(" ", "_")}.exe'
            if os.path.exists(exe_path):
                print(f"✅ Executable created: {exe_path}")
                
                # Try to run the executable
                print("Testing executable...")
                try:
                    run_result = subprocess.run([exe_path], capture_output=True, text=True, timeout=30)
                    if run_result.returncode == 0:
                        print("✅ Executable runs successfully!")
                        print("Output:", run_result.stdout)
                    else:
                        print("⚠️ Executable created but failed to run")
                        print("STDERR:", run_result.stderr)
                except Exception as e:
                    print(f"⚠️ Could not test executable: {e}")
            
            return True
        else:
            print(f"❌ FAILED: {test_name}")
            
            # Show first part of error
            stderr_lines = result.stderr.split('\n')
            print("Key errors:")
            for line in stderr_lines:
                if 'error' in line.lower() or 'traceback' in line.lower() or 'unicodedecodeerror' in line.lower():
                    print(f"  {line}")
            
            return False
            
    except subprocess.TimeoutExpired:
        print(f"⏰ TIMEOUT: {test_name}")
        return False
    except Exception as e:
        print(f"❌ EXCEPTION: {test_name} - {e}")
        return False
    finally:
        # Cleanup
        try:
            os.remove(test_file)
            spec_file = test_file.replace('.py', '.spec')
            if os.path.exists(spec_file):
                os.remove(spec_file)
        except:
            pass

def main():
    print("PROGRESSIVE PYINSTALLER BUILD TEST")
    print("This will test packages incrementally to find the problematic one")
    
    # Test progression - start small and add packages
    test_cases = [
        ([], "Baseline (no packages)"),
        (['tkinter'], "Just Tkinter"),
        (['numpy'], "Just NumPy"),
        (['pandas'], "Just Pandas"),
        (['matplotlib'], "Just Matplotlib"),
        (['PIL'], "Just PIL/Pillow"),
        (['openpyxl'], "Just OpenPyXL"),
        (['numpy', 'pandas'], "NumPy + Pandas"),
        (['matplotlib', 'numpy'], "Matplotlib + NumPy"),
        (['pandas', 'openpyxl'], "Pandas + OpenPyXL"),
        (['numpy', 'pandas', 'matplotlib'], "Core Scientific Stack"),
        (['numpy', 'pandas', 'matplotlib', 'PIL'], "Core + PIL"),
        (['numpy', 'pandas', 'matplotlib', 'PIL', 'openpyxl'], "Core + Office"),
        (['pptx'], "Just python-pptx (likely culprit)"),
        (['numpy', 'pandas', 'matplotlib', 'PIL', 'openpyxl', 'pptx'], "Everything"),
    ]
    
    results = {}
    
    print(f"Will run {len(test_cases)} tests...")
    
    for packages, test_name in test_cases:
        success = test_build_with_packages(packages, test_name)
        results[test_name] = success
        
        if not success:
            print(f"\n🔍 FOUND PROBLEM: {test_name}")
            print("The issue occurs when including these packages:", packages)
            
            if len(packages) > 1:
                print("\nTesting individual packages from this combination...")
                for pkg in packages:
                    individual_success = test_build_with_packages([pkg], f"Individual {pkg}")
                    results[f"Individual {pkg}"] = individual_success
            
            break
        
        time.sleep(2)  # Brief pause between tests
    
    # Summary
    print("\n" + "="*60)
    print("TEST RESULTS SUMMARY")
    print("="*60)
    
    for test_name, success in results.items():
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"{status} {test_name}")
    
    print(f"\nTotal tests run: {len(results)}")
    successful_tests = sum(results.values())
    print(f"Successful: {successful_tests}")
    print(f"Failed: {len(results) - successful_tests}")

if __name__ == "__main__":
    main()