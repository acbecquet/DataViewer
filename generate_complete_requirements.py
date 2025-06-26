#!/usr/bin/env python3
"""
Comprehensive requirements generator.
Scans your Python files for imports and matches with installed packages.
"""
import subprocess
import sys
import os
import glob
import ast
import re
from typing import Set, Dict, List

def get_all_installed_packages() -> Dict[str, str]:
    """Get all installed packages with versions."""
    try:
        result = subprocess.run([sys.executable, '-m', 'pip', 'list'], 
                              capture_output=True, text=True, check=True)
        
        packages = {}
        lines = result.stdout.strip().split('\n')[2:]  # Skip header
        
        for line in lines:
            if line.strip():
                parts = line.split()
                if len(parts) >= 2:
                    name = parts[0].strip()
                    version = parts[1].strip()
                    packages[name.lower()] = f"{name}=={version}"
        
        print(f"Found {len(packages)} installed packages")
        return packages
        
    except Exception as e:
        print(f"Error getting installed packages: {e}")
        return {}

def scan_python_files_for_imports() -> Set[str]:
    """Scan all Python files in current directory for import statements."""
    imports = set()
    
    # Get all Python files
    python_files = glob.glob("*.py")
    
    print(f"Scanning {len(python_files)} Python files for imports...")
    
    for file_path in python_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Parse the AST to find imports
            try:
                tree = ast.parse(content)
                
                for node in ast.walk(tree):
                    if isinstance(node, ast.Import):
                        for alias in node.names:
                            imports.add(alias.name.split('.')[0])
                    elif isinstance(node, ast.ImportFrom):
                        if node.module:
                            imports.add(node.module.split('.')[0])
            
            except SyntaxError:
                # If AST parsing fails, try regex as fallback
                import_patterns = [
                    r'^import\s+([a-zA-Z_][a-zA-Z0-9_]*)',
                    r'^from\s+([a-zA-Z_][a-zA-Z0-9_]*)\s+import'
                ]
                
                for line in content.split('\n'):
                    line = line.strip()
                    for pattern in import_patterns:
                        match = re.match(pattern, line)
                        if match:
                            imports.add(match.group(1))
        
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
    
    # Filter out built-in modules
    builtin_modules = {
        'os', 'sys', 'json', 're', 'time', 'datetime', 'threading', 'queue', 
        'tempfile', 'uuid', 'copy', 'traceback', 'logging', 'shutil', 'subprocess',
        'collections', 'itertools', 'functools', 'pathlib', 'typing', 'enum',
        'abc', 'warnings', 'math', 'random', 'string', 'io', 'csv', 'sqlite3',
        'pickle', 'base64', 'hashlib', 'hmac', 'socket', 'urllib', 'http',
        'email', 'smtplib', 'ftplib', 'zipfile', 'tarfile', 'gzip', 'bz2'
    }
    
    external_imports = imports - builtin_modules
    print(f"Found {len(external_imports)} external package imports")
    print(f"External imports: {sorted(external_imports)}")
    
    return external_imports

def get_package_mapping() -> Dict[str, str]:
    """Map import names to package names (some packages have different import names)."""
    return {
        'cv2': 'opencv-python',
        'PIL': 'pillow',
        'sklearn': 'scikit-learn',
        'yaml': 'pyyaml',
        'dateutil': 'python-dateutil',
        'win32api': 'pywin32',
        'win32con': 'pywin32',
        'win32gui': 'pywin32',
        'pywintypes': 'pywin32',
        'pptx': 'python-pptx',
        'xlsxwriter': 'XlsxWriter',
        'openpyxl': 'openpyxl',
        'tkinter': None,  # Built into Python
        'tkintertable': 'tkintertable',
        'matplotlib': 'matplotlib',
        'numpy': 'numpy',
        'pandas': 'pandas',
        'requests': 'requests',
        'psutil': 'psutil',
        'sqlalchemy': 'sqlalchemy',
        'scipy': 'scipy',
        'seaborn': 'seaborn',
        'plotly': 'plotly',
        'packaging': 'packaging',
        'pkg_resources': 'setuptools',
    }

def generate_requirements(installed_packages: Dict[str, str], 
                        imports: Set[str]) -> tuple[List[str], List[str]]:
    """Generate both regular and Docker requirements."""
    
    package_mapping = get_package_mapping()
    
    # Get required packages
    required_packages = set()
    
    for import_name in imports:
        # Check if we have a mapping for this import
        if import_name in package_mapping:
            pkg_name = package_mapping[import_name]
            if pkg_name:  # None means it's built-in
                required_packages.add(pkg_name.lower())
        else:
            # Assume import name is the same as package name
            required_packages.add(import_name.lower())
    
    # Add packages that we know are needed but might not be directly imported
    always_needed = {
        'setuptools',  # Often needed for packaging
        'wheel',       # For building packages
        'packaging',   # Version parsing
    }
    
    required_packages.update(always_needed)
    
    # Match with installed packages
    all_requirements = []
    docker_requirements = []
    
    # Windows-specific packages to exclude from Docker
    windows_only = {
        'pywin32', 'pywin32-ctypes', 'pefile', 'pyinstaller', 
        'pyinstaller-hooks-contrib', 'altgraph'
    }
    
    # Development-only packages to exclude from Docker
    dev_only = {
        'pyinstaller', 'pyinstaller-hooks-contrib', 'altgraph',
        'setuptools', 'wheel', 'pip'
    }
    
    for pkg_name in sorted(required_packages):
        found_package = None
        
        # Look for exact match first
        if pkg_name in installed_packages:
            found_package = installed_packages[pkg_name]
        else:
            # Look for partial matches
            for installed_name, version_string in installed_packages.items():
                if pkg_name in installed_name or installed_name in pkg_name:
                    found_package = version_string
                    break
        
        if found_package:
            all_requirements.append(found_package)
            
            # Add to Docker requirements if not Windows/dev-only
            actual_pkg_name = found_package.split('==')[0].lower()
            if not any(win_pkg in actual_pkg_name for win_pkg in windows_only):
                if not any(dev_pkg in actual_pkg_name for dev_pkg in dev_only):
                    # For Docker, use headless version of opencv if available
                    if 'opencv-python' in actual_pkg_name:
                        docker_requirements.append(found_package.replace('opencv-python', 'opencv-python-headless'))
                    else:
                        docker_requirements.append(found_package)
        else:
            print(f"Warning: Required package '{pkg_name}' not found in installed packages")
    
    return sorted(all_requirements), sorted(docker_requirements)

def write_requirements_files(all_reqs: List[str], docker_reqs: List[str]):
    """Write both requirements files safely."""
    
    print("\nWriting requirements files...")
    
    # Write complete requirements.txt
    try:
        with open('requirements.txt', 'w', encoding='utf-8', newline='\n') as f:
            f.write("# Complete requirements for local development\n")
            f.write("# Generated automatically\n\n")
            for req in all_reqs:
                f.write(req + '\n')
        print(f"✓ Created requirements.txt with {len(all_reqs)} packages")
    except Exception as e:
        print(f"Error writing requirements.txt: {e}")
    
    # Write Docker requirements
    try:
        with open('requirements-docker.txt', 'w', encoding='utf-8', newline='\n') as f:
            f.write("# Docker-optimized requirements (Linux, runtime only)\n")
            f.write("# Generated automatically\n\n")
            for req in docker_reqs:
                f.write(req + '\n')
        print(f"✓ Created requirements-docker.txt with {len(docker_reqs)} packages")
    except Exception as e:
        print(f"Error writing requirements-docker.txt: {e}")

def main():
    print("Comprehensive Requirements Generator")
    print("=" * 60)
    
    print("\nStep 1: Getting installed packages...")
    installed_packages = get_all_installed_packages()
    
    print("\nStep 2: Scanning Python files for imports...")
    imports = scan_python_files_for_imports()
    
    print("\nStep 3: Generating requirements...")
    all_reqs, docker_reqs = generate_requirements(installed_packages, imports)
    
    print("\nStep 4: Writing files...")
    write_requirements_files(all_reqs, docker_reqs)
    
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"✓ Scanned {len(glob.glob('*.py'))} Python files")
    print(f"✓ Found {len(imports)} external imports")
    print(f"✓ Generated {len(all_reqs)} total requirements")
    print(f"✓ Generated {len(docker_reqs)} Docker requirements")
    
    print(f"\nExcluded from Docker ({len(all_reqs) - len(docker_reqs)} packages):")
    excluded = set(all_reqs) - set(docker_reqs)
    for pkg in sorted(excluded):
        print(f"  - {pkg}")
    
    print(f"\nFiles created:")
    print(f"  - requirements.txt (complete)")
    print(f"  - requirements-docker.txt (Docker optimized)")
    
    print(f"\nNext steps:")
    print(f"  1. Review the generated files")
    print(f"  2. Run: docker-build.bat")
    print(f"  3. Test: docker-run.bat")

if __name__ == "__main__":
    main()