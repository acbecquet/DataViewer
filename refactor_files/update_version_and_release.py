#!/usr/bin/env python3
"""
Complete release workflow script
Handles version incrementing, building, and GitHub release creation
"""
import re
import subprocess
import sys
import os
import json
from datetime import datetime

def debug_print(message):
    """Print debug messages with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"RELEASE_DEBUG [{timestamp}]: {message}")

def get_current_version():
    """Get current version from setup.py"""
    debug_print("Reading current version from setup.py...")
    
    try:
        with open('setup.py', 'r') as f:
            content = f.read()
        
        # Extract version from setup.py
        version_match = re.search(r'version="([^"]+)"', content)
        if version_match:
            current_version = version_match.group(1)
            debug_print(f"Current version: {current_version}")
            return current_version
        else:
            debug_print("ERROR: Could not find version in setup.py")
            return None
    except Exception as e:
        debug_print(f"ERROR: Could not read setup.py: {e}")
        return None

def increment_version(current_version, increment_type='patch'):
    """Increment version number"""
    debug_print(f"Incrementing version: {current_version} ({increment_type})")
    
    try:
        parts = current_version.split('.')
        major, minor, patch = int(parts[0]), int(parts[1]), int(parts[2])
        
        if increment_type == 'major':
            major += 1
            minor = 0
            patch = 0
        elif increment_type == 'minor':
            minor += 1
            patch = 0
        else:  # patch
            patch += 1
        
        new_version = f"{major}.{minor}.{patch}"
        debug_print(f"New version: {new_version}")
        return new_version
        
    except Exception as e:
        debug_print(f"ERROR: Could not increment version: {e}")
        return None

def update_version_in_files(new_version):
    """Update version in all relevant files"""
    debug_print(f"Updating version to {new_version} in all files...")
    
    files_to_update = {
        'setup.py': [
            (r'version="[\d\.]+"', f'version="{new_version}"')
        ],
        'installer_script.iss': [
            (r'AppVersion=[\d\.]+', f'AppVersion={new_version}'),
            (r'OutputBaseFilename=TestingGUI_Setup_v[\d\.]+', f'OutputBaseFilename=TestingGUI_Setup_v{new_version}')
        ],
        'build_exe.py': [
            (r'--name=TestingGUI', f'--name=TestingGUI_v{new_version}')  # Optional: version in exe name
        ]
    }
    
    # Also update version in main_gui.py if it has a version constant
    try:
        with open('main_gui.py', 'r') as f:
            content = f.read()
        
        if 'VERSION' in content or 'version' in content:
            files_to_update['main_gui.py'] = [
                (r'VERSION\s*=\s*["\'][\d\.]+["\']', f'VERSION = "{new_version}"'),
                (r'version\s*=\s*["\'][\d\.]+["\']', f'version = "{new_version}"')
            ]
    except:
        pass
    
    updated_files = []
    for filename, replacements in files_to_update.items():
        try:
            if not os.path.exists(filename):
                debug_print(f"SKIP: File not found: {filename}")
                continue
                
            with open(filename, 'r') as f:
                content = f.read()
            
            original_content = content
            for pattern, replacement in replacements:
                content = re.sub(pattern, replacement, content)
            
            if content != original_content:
                with open(filename, 'w') as f:
                    f.write(content)
                updated_files.append(filename)
                debug_print(f"UPDATED: {filename}")
            else:
                debug_print(f"NO CHANGE: {filename}")
                
        except Exception as e:
            debug_print(f"ERROR updating {filename}: {e}")
    
    return updated_files

def build_installer():
    """Build the installer with new version"""
    debug_print("Building installer...")
    
    try:
        # Run the build process
        result = subprocess.run(['build_installer.bat'], 
                              shell=True, 
                              capture_output=True, 
                              text=True,
                              timeout=300)  # 5 minute timeout
        
        if result.returncode == 0:
            debug_print("Installer built successfully")
            debug_print(f"Build output: {result.stdout[-200:]}")  # Last 200 chars
            return True
        else:
            debug_print(f"Build failed with return code: {result.returncode}")
            debug_print(f"Build error: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        debug_print("ERROR: Build timed out after 5 minutes")
        return False
    except Exception as e:
        debug_print(f"ERROR: Build process failed: {e}")
        return False

def create_github_release(version, release_notes=""):
    """Create GitHub release using GitHub CLI"""
    debug_print(f"Creating GitHub release for v{version}...")
    
    # Check if GitHub CLI is installed
    try:
        subprocess.run(['gh', '--version'], 
                      capture_output=True, 
                      check=True)
        debug_print("GitHub CLI found")
    except (subprocess.CalledProcessError, FileNotFoundError):
        debug_print("ERROR: GitHub CLI not found. Install from: https://cli.github.com/")
        return False
    
    try:
        # Find the installer file
        installer_pattern = f"installer_output/TestingGUI_Setup_v{version}.exe"
        if not os.path.exists(installer_pattern):
            # Try alternative patterns
            alternatives = [
                f"installer_output/TestingGUI_Setup_{version}.exe",
                f"dist/TestingGUI_Setup_v{version}.exe",
                "installer_output/TestingGUI_Setup.exe"
            ]
            
            installer_path = None
            for alt in alternatives:
                if os.path.exists(alt):
                    installer_path = alt
                    break
            
            if not installer_path:
                debug_print(f"ERROR: Installer not found. Looked for: {installer_pattern}")
                return False
        else:
            installer_path = installer_pattern
        
        debug_print(f"Found installer: {installer_path}")
        
        # Create the release
        tag_name = f"v{version}"
        release_title = f"TestingGUI v{version}"
        
        if not release_notes:
            release_notes = f"Release v{version}\n\nChanges:\n- Bug fixes and improvements"
        
        cmd = [
            'gh', 'release', 'create',
            tag_name,
            installer_path,
            '--title', release_title,
            '--notes', release_notes
        ]
        
        debug_print(f"Running: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, 
                              capture_output=True, 
                              text=True,
                              check=True)
        
        debug_print("GitHub release created successfully!")
        debug_print(f"Release URL: {result.stdout.strip()}")
        return True
        
    except subprocess.CalledProcessError as e:
        debug_print(f"ERROR: GitHub release failed: {e.stderr}")
        return False
    except Exception as e:
        debug_print(f"ERROR: Release creation failed: {e}")
        return False

def commit_version_changes(version, updated_files):
    """Commit version changes to git"""
    debug_print(f"Committing version changes...")
    
    try:
        # Add updated files
        for file in updated_files:
            subprocess.run(['git', 'add', file], check=True)
        
        # Commit
        commit_message = f"Bump version to {version}"
        subprocess.run(['git', 'commit', '-m', commit_message], check=True)
        
        # Push
        subprocess.run(['git', 'push', 'origin', 'main'], check=True)
        
        debug_print("Version changes committed and pushed")
        return True
        
    except subprocess.CalledProcessError as e:
        debug_print(f"ERROR: Git operations failed: {e}")
        return False

def main():
    """Main release workflow"""
    print("=" * 60)
    print("TestingGUI Release Workflow")
    print("=" * 60)
    
    # Get current version
    current_version = get_current_version()
    if not current_version:
        print("FAILED: Could not determine current version")
        return False
    
    # Ask user for increment type
    print(f"\nCurrent version: {current_version}")
    print("How do you want to increment the version?")
    print("1. Patch (3.0.0 -> 3.0.1) - Bug fixes")
    print("2. Minor (3.0.0 -> 3.1.0) - New features")
    print("3. Major (3.0.0 -> 4.0.0) - Breaking changes")
    print("4. Custom version")
    
    choice = input("Enter choice (1-4): ").strip()
    
    if choice == '1':
        new_version = increment_version(current_version, 'patch')
    elif choice == '2':
        new_version = increment_version(current_version, 'minor')
    elif choice == '3':
        new_version = increment_version(current_version, 'major')
    elif choice == '4':
        new_version = input("Enter custom version (e.g., 3.0.1): ").strip()
    else:
        print("Invalid choice")
        return False
    
    if not new_version:
        print("FAILED: Could not determine new version")
        return False
    
    # Get release notes
    release_notes = input(f"\nEnter release notes for v{new_version} (optional): ").strip()
    
    # Confirm
    print(f"\nRelease Summary:")
    print(f"  Current Version: {current_version}")
    print(f"  New Version: {new_version}")
    print(f"  Release Notes: {release_notes or 'Default notes'}")
    
    confirm = input("\nProceed with release? (y/N): ").lower().strip()
    if confirm != 'y':
        print("Release cancelled")
        return False
    
    # Execute release workflow
    print("\n" + "=" * 40)
    print("STARTING RELEASE WORKFLOW")
    print("=" * 40)
    
    # Step 1: Update version in files
    updated_files = update_version_in_files(new_version)
    if not updated_files:
        print("FAILED: No files were updated")
        return False
    
    # Step 2: Build installer
    if not build_installer():
        print("FAILED: Could not build installer")
        return False
    
    # Step 3: Commit changes
    if not commit_version_changes(new_version, updated_files):
        print("WARNING: Could not commit changes (continuing anyway)")
    
    # Step 4: Create GitHub release
    if not create_github_release(new_version, release_notes):
        print("FAILED: Could not create GitHub release")
        return False
    
    print("\n" + "=" * 60)
    print("SUCCESS: Release workflow completed!")
    print(f"Version {new_version} has been released")
    print("Users will be notified of the update automatically")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)