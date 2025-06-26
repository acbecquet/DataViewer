#!/usr/bin/env python3
"""Check for missing imports in all Python files."""
import os
import glob
import ast
import importlib.util

def check_file_imports(filepath):
    """Check if all imports in a file are available."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        tree = ast.parse(content)
        imports = []
        
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    imports.append(alias.name)
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    imports.append(node.module)
        
        missing = []
        for imp in imports:
            # Skip relative imports and built-ins
            if imp.startswith('.') or imp in ['sys', 'os', 'json', 're', 'time', 'threading', 'queue', 'tempfile', 'uuid', 'copy', 'traceback']:
                continue
            
            try:
                importlib.import_module(imp.split('.')[0])
            except ImportError:
                missing.append(imp)
        
        if missing:
            print(f"\n{filepath}:")
            for imp in missing:
                print(f"  ✗ Missing: {imp}")
        else:
            print(f"✓ {filepath}")
            
    except Exception as e:
        print(f"Error checking {filepath}: {e}")

def main():
    """Check all Python files for missing imports."""
    print("Checking imports in all Python files...")
    
    python_files = glob.glob("*.py")
    for file in python_files:
        check_file_imports(file)
    
    print("\nDone checking imports.")

if __name__ == "__main__":
    main()