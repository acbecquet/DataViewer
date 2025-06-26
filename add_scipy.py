#!/usr/bin/env python3
"""Add scipy to requirements-docker.txt"""

def add_scipy_to_requirements():
    """Add scipy and any other missing scientific packages."""
    
    # Read current requirements
    try:
        with open('requirements-docker.txt', 'r', encoding='utf-8') as f:
            current_reqs = f.read().strip().split('\n')
    except FileNotFoundError:
        print("requirements-docker.txt not found!")
        return
    
    # Check if scipy is already there
    has_scipy = any('scipy' in req.lower() for req in current_reqs)
    
    if not has_scipy:
        print("Adding scipy to requirements...")
        current_reqs.append("scipy>=1.10.0")
    
    # Add any other commonly missing scientific packages
    missing_packages = []
    
    # Check for other potential missing packages
    package_checks = {
        'scikit-learn': 'scikit-learn>=1.0.0',
        'seaborn': 'seaborn>=0.11.0',
        'statsmodels': 'statsmodels>=0.13.0'
    }
    
    for pkg_name, pkg_req in package_checks.items():
        has_pkg = any(pkg_name in req.lower() for req in current_reqs)
        if not has_pkg:
            missing_packages.append(pkg_req)
    
    # Write updated requirements
    with open('requirements-docker.txt', 'w', encoding='utf-8', newline='\n') as f:
        for req in current_reqs:
            if req.strip():  # Skip empty lines
                f.write(req.strip() + '\n')
    
    print(f"✓ Updated requirements-docker.txt")
    print(f"✓ Added scipy")
    if missing_packages:
        print(f"✓ Consider adding these if needed: {missing_packages}")
    
    # Show final requirements count
    print(f"✓ Total packages: {len([r for r in current_reqs if r.strip()])}")

if __name__ == "__main__":
    add_scipy_to_requirements()