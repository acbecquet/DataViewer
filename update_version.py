# update_version.py - Run this before building installer
def update_version():
    import re
    
    new_version = input("Enter new version (e.g., 3.1.0): ").strip()
    
    if not re.match(r'^\d+\.\d+\.\d+$', new_version):
        print("Invalid version format. Use X.Y.Z (e.g., 3.1.0)")
        return
    
    version_parts = new_version.split('.')
    
    files_to_update = {
        'setup.py': [
            (r'version="[\d\.]+"', f'version="{new_version}"')
        ],
        'version_info.txt': [
            (r'filevers=\(\d+, \d+, \d+, \d+\)', f'filevers=({version_parts[0]}, {version_parts[1]}, {version_parts[2]}, 0)'),
            (r'prodvers=\(\d+, \d+, \d+, \d+\)', f'prodvers=({version_parts[0]}, {version_parts[1]}, {version_parts[2]}, 0)'),
            (r'FileVersion.*', f'StringStruct(u\'FileVersion\', u\'{new_version}.0\'),'),
            (r'ProductVersion.*', f'StringStruct(u\'ProductVersion\', u\'{new_version}.0\')])'),
        ],
        'testing_gui_installer.iss': [
            (r'#define MyAppVersion "[\d\.]+"', f'#define MyAppVersion "{new_version}"')
        ]
    }
    
    for filename, replacements in files_to_update.items():
        try:
            with open(filename, 'r') as f:
                content = f.read()
            
            for pattern, replacement in replacements:
                content = re.sub(pattern, replacement, content)
            
            with open(filename, 'w') as f:
                f.write(content)
            
            print(f"✓ Updated {filename}")
            
        except FileNotFoundError:
            print(f"⚠ File not found: {filename}")
    
    print(f"\n✓ Version updated to {new_version}")
    print("Next steps:")
    print("1. Test your application thoroughly")  
    print("2. Run: build_installer.bat")
    print("3. Test the installer")
    print("4. Distribute to users")

if __name__ == "__main__":
    update_version()