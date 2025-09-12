#!/usr/bin/env python3
"""
Setup script to create missing files for professional installer
Run this first to prepare your project for installer creation
"""

import os
from PIL import Image, ImageDraw, ImageFont

def create_directories():
    """Create necessary directories"""
    dirs = ['resources', 'installer_output']
    for dir_name in dirs:
        os.makedirs(dir_name, exist_ok=True)
        print(f"✓ Created directory: {dir_name}")

def create_basic_icon():
    """Create a basic app icon if none exists"""
    icon_path = 'resources/icon.ico'

    if os.path.exists(icon_path):
        print(f"✓ Icon already exists: {icon_path}")
        return

    try:
        # Create a simple icon with app initials
        size = 256
        img = Image.new('RGBA', (size, size), (0, 100, 200, 255))
        draw = ImageDraw.Draw(img)

        # Draw a simple design
        margin = 20
        draw.rounded_rectangle(
            [margin, margin, size-margin, size-margin],
            radius=30,
            fill=(255, 255, 255, 255)
        )

        # Add text
        try:
            font_size = 80
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        text = "TG"  # Testing GUI
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        text_x = (size - text_width) // 2
        text_y = (size - text_height) // 2

        draw.text((text_x, text_y), text, fill=(0, 100, 200, 255), font=font)

        # Save as ICO
        img.save(icon_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print(f"✓ Created basic icon: {icon_path}")

    except Exception as e:
        print(f"⚠ Could not create icon: {e}")
        print("  You can create your own icon.ico file in the resources/ folder")

def create_license_file():
    """Create a basic MIT license file"""
    license_path = 'LICENSE.txt'

    if os.path.exists(license_path):
        print(f"✓ License already exists: {license_path}")
        return

    license_text = '''MIT License

Copyright (c) 2025 Charlie Becquet

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

    with open(license_path, 'w') as f:
        f.write(license_text)
    print(f"✓ Created license file: {license_path}")

def create_readme_file():
    """Create a user README file"""
    readme_path = 'README.txt'

    if os.path.exists(readme_path):
        print(f"✓ README already exists: {readme_path}")
        return

    readme_text = '''Standardized Testing GUI v3.0.0

Welcome to the Standardized Testing GUI - Professional Data Analysis Software

SYSTEM REQUIREMENTS:
- Windows 10 or later (64-bit)
- 4GB RAM minimum (8GB recommended)
- 500MB available disk space
- Display resolution: 1024x768 minimum

FEATURES:
- Excel data processing and analysis
- Advanced plotting and visualization
- Report generation (Excel and PowerPoint)
- Trend analysis tools
- Viscosity calculations
- Professional data export

GETTING STARTED:
1. Launch the application from the Start menu
2. Load your Excel data files
3. Select the analysis type
4. Generate reports and plots

FILE SUPPORT:
- Excel files (.xlsx, .xls)
- VAP3 project files (.vap3)
- Image files for documentation

SUPPORT:
For technical support, documentation, and updates:
- Website: https://your-website.com
- Email: support@your-website.com

Copyright © 2025 Charlie Becquet. All rights reserved.
'''

    with open(readme_path, 'w') as f:
        f.write(readme_text)
    print(f"✓ Created README file: {readme_path}")

def main():
    """Main setup function"""
    print("Setting up files for professional installer...")
    print("=" * 50)

    create_directories()
    create_basic_icon()
    create_license_file()
    create_readme_file()

    print("\n" + "=" * 50)
    print("✓ Setup complete!")
    print("\nNext steps:")
    print("1. Install Inno Setup 6 from: https://jrsoftware.org/isdl.php")
    print("2. Run: build_installer.bat")
    print("3. Your professional installer will be created!")
    print("\nOptional improvements:")
    print("- Replace resources/icon.ico with your custom icon")
    print("- Customize LICENSE.txt and README.txt")
    print("- Add installer background images to resources/")

if __name__ == "__main__":
    main()
