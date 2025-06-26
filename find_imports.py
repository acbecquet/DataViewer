import os
import glob

def find_imports():
    """Find all import statements in Python files."""
    python_files = glob.glob("*.py")
    
    for file in python_files:
        try:
            with open(file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
            for i, line in enumerate(lines, 1):
                if 'import requests' in line or 'from requests' in line:
                    print(f"{file}:{i}: {line.strip()}")
        except Exception as e:
            print(f"Error reading {file}: {e}")

if __name__ == "__main__":
    find_imports()