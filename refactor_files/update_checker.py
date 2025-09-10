#!/usr/bin/env python3
"""
update_checker.py - Updated with proper GitHub configuration
"""
import requests
import json
import os
import sys
import subprocess
import tempfile
from packaging import version
from typing import Optional, Dict, Any
import time

class UpdateChecker:
    """Check for and install updates from GitHub"""
    
    def __init__(self, current_version: str = "3.0.0"):
        self.current_version = current_version
        
        # THIS IS WHAT YOU CONFIGURE:
        # Replace 'acbecquet/DataViewer' with your actual GitHub repo
        self.github_repo = "acbecquet/DataViewer"  # Format: username/repository
        
        # GitHub API endpoints (don't change these)
        self.github_api_base = "https://api.github.com"
        self.releases_endpoint = f"{self.github_api_base}/repos/{self.github_repo}/releases"
        
        # You can leave this empty for GitHub-only updates
        self.update_server = ""  # Not used when using GitHub
        
        self.debug_mode = True
        
        print(f"UPDATE_DEBUG: Initialized for repo: {self.github_repo}")
        print(f"UPDATE_DEBUG: Current version: {self.current_version}")
        print(f"UPDATE_DEBUG: Releases API: {self.releases_endpoint}")
    
    def debug_print(self, message: str):
        """Print debug information with timestamp"""
        if self.debug_mode:
            timestamp = time.strftime("%H:%M:%S")
            print(f"UPDATE_DEBUG [{timestamp}]: {message}")
    
    def check_for_updates(self) -> Optional[Dict[str, Any]]:
        """Check GitHub releases for updates"""
        self.debug_print("Checking GitHub releases for updates...")
        
        try:
            # Get latest release from GitHub API
            latest_url = f"{self.releases_endpoint}/latest"
            self.debug_print(f"Requesting: {latest_url}")
            
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': f'TestingGUI-UpdateChecker/{self.current_version}'
            }
            
            response = requests.get(latest_url, headers=headers, timeout=10)
            
            self.debug_print(f"Response status: {response.status_code}")
            
            if response.status_code == 404:
                self.debug_print("No releases found (404)")
                return {'update_available': False, 'reason': 'No releases found'}
            
            response.raise_for_status()
            
            release_data = response.json()
            
            # Extract version from tag (remove 'v' prefix if present)
            tag_name = release_data.get('tag_name', '')
            latest_version = tag_name.lstrip('v')
            
            self.debug_print(f"Latest release tag: {tag_name}")
            self.debug_print(f"Latest version: {latest_version}")
            self.debug_print(f"Current version: {self.current_version}")
            
            # Compare versions using packaging library
            try:
                if version.parse(latest_version) > version.parse(self.current_version):
                    self.debug_print("Update available!")
                    
                    # Find the installer asset
                    installer_asset = None
                    for asset in release_data.get('assets', []):
                        asset_name = asset.get('name', '').lower()
                        if (asset_name.endswith('.exe') and 
                            ('setup' in asset_name or 'install' in asset_name)):
                            installer_asset = asset
                            break
                    
                    if not installer_asset:
                        self.debug_print("WARNING: No installer found in release assets")
                        # List available assets for debugging
                        assets = [a.get('name') for a in release_data.get('assets', [])]
                        self.debug_print(f"Available assets: {assets}")
                    
                    return {
                        'update_available': True,
                        'latest_version': latest_version,
                        'current_version': self.current_version,
                        'release_notes': release_data.get('body', ''),
                        'download_url': installer_asset.get('browser_download_url') if installer_asset else None,
                        'published_at': release_data.get('published_at'),
                        'installer_name': installer_asset.get('name') if installer_asset else None,
                        'installer_size': installer_asset.get('size') if installer_asset else None
                    }
                else:
                    self.debug_print("No update available (version comparison)")
                    return {'update_available': False}
                    
            except Exception as e:
                self.debug_print(f"Error comparing versions: {e}")
                return None
                
        except requests.RequestException as e:
            self.debug_print(f"Network error checking for updates: {e}")
            return None
        except Exception as e:
            self.debug_print(f"Error checking for updates: {e}")
            return None
    
    def download_update(self, download_url: str, filename: str) -> Optional[str]:
        """Download the update installer from GitHub"""
        self.debug_print(f"Downloading update from GitHub: {download_url}")
        
        try:
            # Create temp directory
            temp_dir = tempfile.mkdtemp(prefix="testinggui_update_")
            file_path = os.path.join(temp_dir, filename)
            
            self.debug_print(f"Downloading to: {file_path}")
            
            # Download with progress tracking
            headers = {
                'Accept': 'application/octet-stream',
                'User-Agent': f'TestingGUI-UpdateChecker/{self.current_version}'
            }
            
            response = requests.get(download_url, 
                                  headers=headers,
                                  stream=True, 
                                  timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            self.debug_print(f"File size: {total_size / (1024*1024):.1f} MB")
            
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        if total_size > 0:
                            progress = (downloaded / total_size) * 100
                            if progress % 10 < 1:  # Log every 10%
                                self.debug_print(f"Download progress: {progress:.1f}%")
            
            self.debug_print(f"Download completed: {file_path}")
            return file_path
            
        except Exception as e:
            self.debug_print(f"Error downloading update: {e}")
            return None
    
    def install_update(self, installer_path: str) -> bool:
        """Install the downloaded update"""
        self.debug_print(f"Installing update from: {installer_path}")
        
        try:
            # Verify file exists
            if not os.path.exists(installer_path):
                self.debug_print("ERROR: Installer file not found")
                return False
            
            # Verify file size (should be substantial)
            size_mb = os.path.getsize(installer_path) / (1024 * 1024)
            self.debug_print(f"Installer size: {size_mb:.1f} MB")
            
            if size_mb < 1:  # Less than 1MB seems wrong
                self.debug_print("WARNING: Installer seems too small")
            
            # Run the installer in silent mode
            self.debug_print("Launching installer...")
            
            # Use different flags depending on installer type
            # Inno Setup supports /SILENT
            subprocess.Popen([installer_path, '/SILENT'], 
                           shell=False, 
                           cwd=os.path.dirname(installer_path))
            
            # Give installer time to start
            time.sleep(2)
            
            # Exit current application to allow update
            self.debug_print("Exiting application for update...")
            sys.exit(0)
            
        except Exception as e:
            self.debug_print(f"Error installing update: {e}")
            return False

# Test function
def test_update_checker():
    """Test the update checker"""
    print("Testing Update Checker")
    print("=" * 40)
    
    # You can test with different versions
    checker = UpdateChecker(current_version="2.9.9")  # Use older version to trigger update
    
    update_info = checker.check_for_updates()
    
    if update_info is None:
        print("❌ Could not check for updates (network/API error)")
        return
    
    print(f"Update check result: {update_info}")
    
    if update_info.get('update_available'):
        print(f"✅ Update available: v{update_info['latest_version']}")
        print(f"📝 Release notes: {update_info['release_notes'][:100]}...")
        
        if update_info.get('download_url'):
            print(f"📦 Download URL: {update_info['download_url']}")
            
            # Uncomment to test download (be careful!)
            # installer_path = checker.download_update(
            #     update_info['download_url'],
            #     update_info['installer_name']
            # )
            # if installer_path:
            #     print(f"✅ Downloaded to: {installer_path}")
        else:
            print("❌ No download URL found")
    else:
        print("ℹ️ No updates available")

if __name__ == "__main__":
    test_update_checker()