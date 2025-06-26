#!/usr/bin/env python3
"""Create Docker files with correct encoding in unencrypted location."""

def create_docker_build_bat():
    content = '''@echo off
echo Building Docker image for Testing GUI...
echo ========================================

REM Check if Docker is running
docker info >nul 2>&1
if errorlevel 1 (
    echo ERROR: Docker is not running. Please start Docker Desktop.
    pause
    exit /b 1
)

REM Build the Docker image
echo Building image...
docker build -t testing-gui:latest .

if errorlevel 1 (
    echo ERROR: Failed to build Docker image.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Docker image built successfully!
echo ========================================
pause
'''
    
    with open('docker-build.bat', 'w', encoding='ascii') as f:
        f.write(content)
    print("✓ Created docker-build.bat")

def create_simple_dockerfile():
    content = '''# Use Python 3.11 slim image
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install system dependencies for GUI
RUN apt-get update && apt-get install -y \\
    python3-tk \\
    libglib2.0-0 \\
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python packages
COPY requirements-docker.txt .
RUN pip install --no-cache-dir -r requirements-docker.txt

# Copy application code
COPY . .

# Install the application
RUN pip install -e .

# Run the application
CMD ["python", "main.py"]
'''
    
    with open('Dockerfile', 'w', encoding='ascii') as f:
        f.write(content)
    print("✓ Created Dockerfile")

def create_requirements_docker():
    content = '''matplotlib==3.9.2
numpy==2.0.0
pandas==2.2.2
openpyxl==3.1.5
pillow==10.4.0
psutil==6.0.0
sqlalchemy==2.0.27
requests>=2.25.0
opencv-python-headless>=4.5.0
python-pptx==0.6.23
XlsxWriter==3.2.0
tkintertable==1.3.3
python-dateutil==2.9.0.post0
pytz==2024.1
packaging>=20.0
'''
    
    with open('requirements-docker.txt', 'w', encoding='ascii') as f:
        f.write(content)
    print("✓ Created requirements-docker.txt")

def create_dockerignore():
    content = '''__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg
venv/
env/
ENV/
.vscode/
.idea/
*.swp
*.swo
*~
.DS_Store
Thumbs.db
.git/
.gitignore
*.log
logs/
*.xlsx
*.xls
*.csv
data/
Dockerfile*
docker-compose*
*.bat
README*.md
docs/
'''
    
    with open('.dockerignore', 'w', encoding='ascii') as f:
        f.write(content)
    print("✓ Created .dockerignore")

def main():
    print("Creating Docker files in current directory...")
    create_docker_build_bat()
    create_simple_dockerfile()
    create_requirements_docker()
    create_dockerignore()
    print("\n" + "="*50)
    print("All Docker files created successfully!")
    print("="*50)
    print("Next steps:")
    print("1. Copy your Python source files to this directory")
    print("2. Run: docker-build.bat")
    print("3. Install VcXsrv for GUI support")

if __name__ == "__main__":
    main()