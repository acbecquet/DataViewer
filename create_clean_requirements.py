#!/usr/bin/env python3
"""Create a clean requirements-docker.txt file with proper encoding."""

def create_clean_requirements():
    # Clean, simple requirements list
    requirements = [
        "matplotlib==3.9.2",
        "numpy==2.0.0", 
        "pandas==2.2.2",
        "openpyxl==3.1.5",
        "pillow==10.4.0",
        "psutil==6.0.0",
        "sqlalchemy==2.0.27",
        "requests>=2.25.0",
        "opencv-python-headless>=4.5.0",
        "python-pptx==0.6.23",
        "XlsxWriter==3.2.0",
        "tkintertable==1.3.3",
        "python-dateutil==2.9.0.post0",
        "pytz==2024.1",
        "packaging>=20.0"
    ]
    
    # Write with explicit UTF-8 encoding
    with open('requirements-docker.txt', 'w', encoding='utf-8', newline='\n') as f:
        for req in requirements:
            f.write(req + '\n')
    
    print("✓ Created clean requirements-docker.txt")
    
    # Verify the file
    with open('requirements-docker.txt', 'r', encoding='utf-8') as f:
        content = f.read()
        print(f"✓ File has {len(content.splitlines())} lines")
        print("✓ First few lines:")
        for i, line in enumerate(content.splitlines()[:5]):
            print(f"  {i+1}: {line}")

if __name__ == "__main__":
    create_clean_requirements()