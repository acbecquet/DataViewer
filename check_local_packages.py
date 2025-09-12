#!/usr/bin/env python3
"""Check what packages are actually needed by scanning imports."""
import subprocess
import sys

def get_installed_packages():
    """Get list of installed packages."""
    result = subprocess.run([sys.executable, '-m', 'pip', 'list'],
                          capture_output=True, text=True)
    return result.stdout

def main():
    print("Currently installed packages:")
    print("=" * 50)
    packages = get_installed_packages()

    # Look for scientific packages
    scientific_packages = ['scipy', 'scikit-learn', 'seaborn', 'statsmodels',
                          'sympy', 'networkx', 'plotly']

    print("Scientific packages found:")
    for line in packages.split('\n'):
        for sci_pkg in scientific_packages:
            if sci_pkg in line.lower():
                print(f"  {line.strip()}")

if __name__ == "__main__":
    main()
