# services/update_service.py
"""
services/update_service.py
Consolidated update service for version management, database updates, and release workflows.
This consolidates update logic from update_checker.py, update_version_and_release.py,
release_workflow.py, and database update functionality from main_gui.py.
"""

import os
import re
import json
import tempfile
import subprocess
import traceback
import threading
import uuid
from typing import Optional, Dict, Any, List, Tuple, Callable, Union
from datetime import datetime, timedelta
from pathlib import Path

# Third party imports
import requests
import pandas as pd


def debug_print(message: str):
    """Debug print function for update operations."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"DEBUG: UpdateService [{timestamp}] - {message}")


def success_print(message: str):
    """Success print function with formatting."""
    print(f"✅ SUCCESS: {message}")


def error_print(message: str):
    """Error print function with formatting."""
    print(f"❌ ERROR: {message}")


def warning_print(message: str):
    """Warning print function with formatting."""
    print(f"⚠️  WARNING: {message}")


class VersionManager:
    """Handles version string parsing and manipulation."""
    
    def __init__(self):
        """Initialize version manager."""
        self.version_pattern = re.compile(r'(\d+)\.(\d+)\.(\d+)')
        
    def parse_version(self, version_string: str) -> Optional[Tuple[int, int, int]]:
        """Parse version string into major, minor, patch tuple."""
        try:
            # Remove any 'v' prefix
            clean_version = version_string.lstrip('v')
            match = self.version_pattern.match(clean_version)
            
            if match:
                return int(match.group(1)), int(match.group(2)), int(match.group(3))
            else:
                debug_print(f"WARNING: Could not parse version string: {version_string}")
                return None
        except Exception as e:
            debug_print(f"ERROR: Version parsing failed: {e}")
            return None
    
    def increment_version(self, current_version: str, increment_type: str = 'patch') -> Optional[str]:
        """Increment version number based on type."""
        try:
            version_tuple = self.parse_version(current_version)
            if not version_tuple:
                return None
            
            major, minor, patch = version_tuple
            
            if increment_type.lower() == 'major':
                major += 1
                minor = 0
                patch = 0
            elif increment_type.lower() == 'minor':
                minor += 1
                patch = 0
            else:  # patch
                patch += 1
            
            new_version = f"{major}.{minor}.{patch}"
            debug_print(f"Incremented version {current_version} -> {new_version} ({increment_type})")
            return new_version
            
        except Exception as e:
            debug_print(f"ERROR: Version increment failed: {e}")
            return None
    
    def compare_versions(self, version1: str, version2: str) -> int:
        """
        Compare two versions.
        Returns: 1 if version1 > version2, -1 if version1 < version2, 0 if equal
        """
        try:
            v1_tuple = self.parse_version(version1)
            v2_tuple = self.parse_version(version2)
            
            if not v1_tuple or not v2_tuple:
                return 0
            
            if v1_tuple > v2_tuple:
                return 1
            elif v1_tuple < v2_tuple:
                return -1
            else:
                return 0
                
        except Exception as e:
            debug_print(f"ERROR: Version comparison failed: {e}")
            return 0
    
    def is_newer_version(self, version1: str, version2: str) -> bool:
        """Check if version1 is newer than version2."""
        return self.compare_versions(version1, version2) > 0


class FileVersionUpdater:
    """Handles updating version strings in source files."""
    
    def __init__(self):
        """Initialize file version updater."""
        self.version_files = {
            'setup.py': [
                (r'version="[\d\.]+"', r'version="{}"'),
                (r"version='[\d\.]+'", r"version='{}'")
            ],
            'main_gui.py': [
                (r'VERSION\s*=\s*["\'][\d\.]+["\']', r'VERSION = "{}"'),
                (r'version\s*=\s*["\'][\d\.]+["\']', r'version = "{}"'),
                (r'__version__\s*=\s*["\'][\d\.]+["\']', r'__version__ = "{}"')
            ],
            'installer_script.iss': [
                (r'AppVersion=[\d\.]+', r'AppVersion={}'),
                (r'OutputBaseFilename=TestingGUI_Setup_v[\d\.]+', r'OutputBaseFilename=TestingGUI_Setup_v{}')
            ],
            'build_exe.py': [
                (r'--name=TestingGUI_v[\d\.]+', r'--name=TestingGUI_v{}')
            ],
            'config.json': [
                (r'"version"\s*:\s*"[\d\.]+"', r'"version": "{}"')
            ]
        }
        
    def update_version_in_files(self, new_version: str) -> List[str]:
        """Update version in all relevant files."""
        debug_print(f"Updating version to {new_version} in all files...")
        
        updated_files = []
        
        for filename, patterns in self.version_files.items():
            if os.path.exists(filename):
                try:
                    success = self._update_file_version(filename, patterns, new_version)
                    if success:
                        updated_files.append(filename)
                        debug_print(f"Updated version in {filename}")
                    else:
                        debug_print(f"WARNING: Could not update version in {filename}")
                except Exception as e:
                    debug_print(f"ERROR: Failed to update {filename}: {e}")
            else:
                debug_print(f"File not found: {filename}")
        
        debug_print(f"Updated version in {len(updated_files)} files")
        return updated_files
    
    def _update_file_version(self, filename: str, patterns: List[Tuple[str, str]], new_version: str) -> bool:
        """Update version in a specific file using regex patterns."""
        try:
            # Read file content
            with open(filename, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_content = content
            updated = False
            
            # Apply all patterns
            for search_pattern, replace_pattern in patterns:
                new_content = re.sub(search_pattern, replace_pattern.format(new_version), content)
                if new_content != content:
                    content = new_content
                    updated = True
            
            # Write back if changed
            if updated:
                # Create backup
                backup_path = f"{filename}.backup"
                with open(backup_path, 'w', encoding='utf-8') as f:
                    f.write(original_content)
                
                # Write updated content
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                debug_print(f"Created backup: {backup_path}")
                return True
            else:
                debug_print(f"No version patterns found in {filename}")
                return False
                
        except Exception as e:
            debug_print(f"ERROR: Failed to update {filename}: {e}")
            return False
    
    def get_current_version_from_setup(self) -> Optional[str]:
        """Get current version from setup.py file."""
        try:
            if not os.path.exists('setup.py'):
                debug_print("setup.py not found")
                return None
            
            with open('setup.py', 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Look for version string
            version_match = re.search(r'version=["\']([^"\']+)["\']', content)
            if version_match:
                version = version_match.group(1)
                debug_print(f"Current version from setup.py: {version}")
                return version
            else:
                debug_print("Could not find version in setup.py")
                return None
                
        except Exception as e:
            debug_print(f"ERROR: Could not read version from setup.py: {e}")
            return None


class DatabaseUpdateManager:
    """Handles database update operations for modified files."""
    
    def __init__(self, database_service=None, file_service=None):
        """Initialize database update manager."""
        self.database_service = database_service
        self.file_service = file_service
        self.modified_files_tracking = {}
        self.update_lock = threading.Lock()
        
    def track_file_modification(self, file_data: Dict[str, Any], modification_type: str = "data_change"):
        """Track that a file has been modified and needs database update."""
        try:
            file_name = file_data.get('file_name', 'unknown')
            
            with self.update_lock:
                self.modified_files_tracking[file_name] = {
                    'file_data': file_data,
                    'modification_type': modification_type,
                    'timestamp': datetime.now(),
                    'update_pending': True
                }
            
            debug_print(f"Tracking modification for file: {file_name} ({modification_type})")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to track file modification: {e}")
    
    def get_modified_files(self) -> List[Dict[str, Any]]:
        """Get list of files that have been modified and need database updates."""
        try:
            with self.update_lock:
                modified_files = []
                for file_name, tracking_data in self.modified_files_tracking.items():
                    if tracking_data.get('update_pending', False):
                        # Add modification metadata
                        file_data = tracking_data['file_data'].copy()
                        file_data['modification_metadata'] = {
                            'type': tracking_data['modification_type'],
                            'timestamp': tracking_data['timestamp'],
                            'tracked_since': tracking_data['timestamp']
                        }
                        modified_files.append(file_data)
                
                debug_print(f"Found {len(modified_files)} modified files needing database update")
                return modified_files
                
        except Exception as e:
            debug_print(f"ERROR: Failed to get modified files: {e}")
            return []
    
    def update_database_with_modified_files(self, modified_files: List[Dict[str, Any]], 
                                          progress_callback: Optional[Callable] = None) -> Tuple[int, List[str]]:
        """
        Update database with all modified files.
        
        Args:
            modified_files: List of file data dictionaries
            progress_callback: Optional callback for progress updates
            
        Returns:
            Tuple of (successful_updates, failed_files)
        """
        debug_print(f"Starting database update for {len(modified_files)} files")
        
        successful_updates = 0
        failed_updates = []
        
        for i, file_data in enumerate(modified_files):
            try:
                if progress_callback:
                    progress = int(((i + 1) / len(modified_files)) * 100)
                    progress_callback(progress)
                
                file_name = file_data.get('file_name', f'file_{i}')
                debug_print(f"Updating database for file: {file_name}")
                
                # Update file in database
                success = self._update_single_file_in_database(file_data)
                
                if success:
                    successful_updates += 1
                    # Mark as updated
                    with self.update_lock:
                        if file_name in self.modified_files_tracking:
                            self.modified_files_tracking[file_name]['update_pending'] = False
                            self.modified_files_tracking[file_name]['last_updated'] = datetime.now()
                else:
                    failed_updates.append(file_name)
                
            except Exception as e:
                file_name = file_data.get('file_name', f'file_{i}')
                debug_print(f"ERROR: Failed to update file {file_name}: {e}")
                failed_updates.append(file_name)
        
        debug_print(f"Database update completed: {successful_updates} successful, {len(failed_updates)} failed")
        return successful_updates, failed_updates
    
    def _update_single_file_in_database(self, file_data: Dict[str, Any]) -> bool:
        """Update a single file in the database."""
        try:
            # Extract file information
            file_name = file_data.get('file_name')
            filtered_sheets = file_data.get('filtered_sheets', {})
            
            # Collect sample notes data
            sample_notes_data = {}
            for sheet_name, sheet_info in filtered_sheets.items():
                header_data = sheet_info.get('header_data')
                if header_data and 'samples' in header_data:
                    sheet_notes = {}
                    for i, sample_data in enumerate(header_data['samples']):
                        sample_notes = sample_data.get('sample_notes', '')
                        if sample_notes.strip():
                            sheet_notes[f"Sample {i+1}"] = sample_notes
                    
                    if sheet_notes:
                        sample_notes_data[sheet_name] = sheet_notes
                        debug_print(f"Collected notes for sheet {sheet_name}: {len(sheet_notes)} samples")
            
            # Create temporary VAP3 file
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
            
            try:
                # Prepare VAP3 data structure
                vap3_data = {
                    'filtered_sheets': filtered_sheets,
                    'sample_notes': sample_notes_data,
                    'sheet_images': file_data.get('sheet_images', {}),
                    'image_crop_states': file_data.get('image_crop_states', {}),
                    'plot_settings': file_data.get('plot_settings', {}),
                    'sample_images': file_data.get('sample_images', {}),
                    'file_metadata': {
                        'original_filename': file_name,
                        'update_timestamp': datetime.now().isoformat(),
                        'version': '3.0'
                    }
                }
                
                # Save to VAP3 format if file service is available
                if self.file_service:
                    success = self.file_service.save_vap3_file(temp_vap3_path, vap3_data)
                    if not success:
                        debug_print(f"WARNING: VAP3 save failed, using JSON fallback")
                        # Fallback to JSON
                        with open(temp_vap3_path, 'w', encoding='utf-8') as f:
                            json.dump(vap3_data, f, indent=2, default=str)
                else:
                    # Fallback to JSON if no file service
                    with open(temp_vap3_path, 'w', encoding='utf-8') as f:
                        json.dump(vap3_data, f, indent=2, default=str)
                
                # Store in database if database service is available
                if self.database_service:
                    success = self.database_service.store_file(temp_vap3_path, file_name, vap3_data)
                    if success:
                        debug_print(f"Successfully updated {file_name} in database")
                        return True
                    else:
                        debug_print(f"ERROR: Database storage failed for {file_name}")
                        return False
                else:
                    debug_print(f"WARNING: No database service available")
                    return False
                
            finally:
                # Clean up temporary file
                try:
                    os.remove(temp_vap3_path)
                except Exception as e:
                    debug_print(f"WARNING: Could not remove temp file {temp_vap3_path}: {e}")
            
        except Exception as e:
            debug_print(f"ERROR: Single file database update failed: {e}")
            traceback.print_exc()
            return False
    
    def clear_modification_flags(self):
        """Clear all modification flags after successful database update."""
        try:
            with self.update_lock:
                for file_name in self.modified_files_tracking:
                    self.modified_files_tracking[file_name]['update_pending'] = False
                    self.modified_files_tracking[file_name]['cleared_at'] = datetime.now()
            
            debug_print("Cleared all modification flags")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to clear modification flags: {e}")


class GitHubReleaseManager:
    """Handles GitHub release creation and management."""
    
    def __init__(self, github_repo: str = "acbecquet/DataViewer"):
        """Initialize GitHub release manager."""
        self.github_repo = github_repo
        self.github_api_base = "https://api.github.com"
        self.releases_endpoint = f"{self.github_api_base}/repos/{self.github_repo}/releases"
        
    def check_github_cli_available(self) -> bool:
        """Check if GitHub CLI is available."""
        try:
            result = subprocess.run(['gh', '--version'], capture_output=True, check=True)
            debug_print("GitHub CLI is available")
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            debug_print("GitHub CLI not found - install from: https://cli.github.com/")
            return False
    
    def create_github_release(self, version: str, installer_path: Optional[str] = None, 
                            release_notes: str = "") -> bool:
        """Create GitHub release with optional installer file."""
        debug_print(f"Creating GitHub release for version {version}")
        
        if not self.check_github_cli_available():
            return False
        
        try:
            tag_name = f"v{version}"
            title = f"DataViewer v{version}"
            
            if not release_notes:
                release_notes = f"Release v{version}\n\n- Bug fixes and improvements"
            
            # Build command
            cmd = [
                'gh', 'release', 'create', tag_name,
                '--title', title,
                '--notes', release_notes
            ]
            
            # Add installer file if provided
            if installer_path and os.path.exists(installer_path):
                cmd.append(installer_path)
                debug_print(f"Including installer file: {installer_path}")
                timeout = 600  # 10 minutes for large files
            else:
                timeout = 60  # 1 minute for release without files
            
            # Execute command
            result = subprocess.run(cmd, capture_output=True, text=True, 
                                  check=True, timeout=timeout)
            
            success_print(f"GitHub release created successfully!")
            debug_print(f"Release URL: {result.stdout.strip()}")
            return True
            
        except subprocess.CalledProcessError as e:
            error_print(f"GitHub release failed: {e.stderr}")
            return False
        except subprocess.TimeoutExpired:
            error_print("GitHub release timed out (large file upload)")
            return False
        except Exception as e:
            error_print(f"Release creation failed: {e}")
            return False
    
    def get_latest_release_info(self) -> Optional[Dict[str, Any]]:
        """Get information about the latest GitHub release."""
        try:
            latest_url = f"{self.releases_endpoint}/latest"
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': f'DataViewer-UpdateService/3.0'
            }
            
            response = requests.get(latest_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            release_data = response.json()
            return {
                'version': release_data.get('tag_name', '').lstrip('v'),
                'release_url': release_data.get('html_url'),
                'download_url': self._extract_download_url(release_data),
                'release_notes': release_data.get('body', ''),
                'published_at': release_data.get('published_at'),
                'prerelease': release_data.get('prerelease', False)
            }
            
        except Exception as e:
            debug_print(f"ERROR: Failed to get latest release info: {e}")
            return None
    
    def _extract_download_url(self, release_data: Dict[str, Any]) -> Optional[str]:
        """Extract download URL from release data."""
        try:
            assets = release_data.get('assets', [])
            for asset in assets:
                name = asset.get('name', '').lower()
                # Look for installer files
                if any(ext in name for ext in ['.exe', '.msi', '.dmg', '.deb', '.rpm']):
                    return asset.get('browser_download_url')
            
            # Fallback to first asset
            if assets:
                return assets[0].get('browser_download_url')
            
            return None
            
        except Exception as e:
            debug_print(f"WARNING: Could not extract download URL: {e}")
            return None


class BuildWorkflowManager:
    """Handles build and release workflow operations."""
    
    def __init__(self):
        """Initialize build workflow manager."""
        self.build_artifacts = ['dist', 'build', '__pycache__']
        self.unwanted_folders = ['test', 'tests', '.pytest_cache', '.git', '.idea']
        
    def cleanup_build_artifacts(self) -> bool:
        """Clean up build artifacts before building."""
        debug_print("Cleaning up build artifacts...")
        
        try:
            import shutil
            
            for artifact in self.build_artifacts:
                if os.path.exists(artifact):
                    if os.path.isdir(artifact):
                        shutil.rmtree(artifact)
                        debug_print(f"Removed directory: {artifact}")
                    else:
                        os.remove(artifact)
                        debug_print(f"Removed file: {artifact}")
            
            debug_print("Build artifact cleanup completed")
            return True
            
        except Exception as e:
            error_print(f"Build artifact cleanup failed: {e}")
            return False
    
    def cleanup_unwanted_folders(self) -> bool:
        """Clean up unwanted folders before building."""
        debug_print("Cleaning up unwanted folders...")
        
        try:
            import shutil
            
            for folder in self.unwanted_folders:
                if os.path.exists(folder):
                    if os.path.isdir(folder):
                        shutil.rmtree(folder)
                        debug_print(f"Removed unwanted directory: {folder}")
            
            debug_print("Unwanted folder cleanup completed")
            return True
            
        except Exception as e:
            error_print(f"Unwanted folder cleanup failed: {e}")
            return False
    
    def build_executable(self) -> bool:
        """Build executable using pyinstaller or similar."""
        debug_print("Building executable...")
        
        try:
            # Look for build script
            build_script = None
            for script_name in ['build_exe.py', 'build.py', 'setup.py']:
                if os.path.exists(script_name):
                    build_script = script_name
                    break
            
            if build_script:
                debug_print(f"Using build script: {build_script}")
                result = subprocess.run(['python', build_script], 
                                      capture_output=True, text=True, timeout=300)
                
                if result.returncode == 0:
                    success_print("Executable build completed")
                    return True
                else:
                    error_print(f"Build script failed: {result.stderr}")
                    return False
            else:
                # Fallback to direct pyinstaller
                debug_print("No build script found, trying direct pyinstaller")
                main_script = None
                for script_name in ['main_gui.py', 'main.py', 'app.py']:
                    if os.path.exists(script_name):
                        main_script = script_name
                        break
                
                if main_script:
                    cmd = ['pyinstaller', '--onefile', '--windowed', main_script]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                    
                    if result.returncode == 0:
                        success_print("Direct pyinstaller build completed")
                        return True
                    else:
                        error_print(f"Pyinstaller failed: {result.stderr}")
                        return False
                else:
                    error_print("No main script found for building")
                    return False
                    
        except subprocess.TimeoutExpired:
            error_print("Build process timed out")
            return False
        except Exception as e:
            error_print(f"Build process failed: {e}")
            return False
    
    def build_installer(self) -> Optional[str]:
        """Build installer and return path to installer file."""
        debug_print("Building installer...")
        
        try:
            # Look for installer script
            installer_script = None
            for script_name in ['installer_script.iss', 'setup.iss', 'installer.iss']:
                if os.path.exists(script_name):
                    installer_script = script_name
                    break
            
            if installer_script:
                debug_print(f"Using installer script: {installer_script}")
                
                # Try to use Inno Setup
                iscc_paths = [
                    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
                    r"C:\Program Files\Inno Setup 6\ISCC.exe",
                    "iscc"  # If in PATH
                ]
                
                iscc_exe = None
                for path in iscc_paths:
                    try:
                        result = subprocess.run([path], capture_output=True)
                        iscc_exe = path
                        break
                    except:
                        continue
                
                if iscc_exe:
                    result = subprocess.run([iscc_exe, installer_script], 
                                          capture_output=True, text=True, timeout=300)
                    
                    if result.returncode == 0:
                        # Find generated installer
                        output_dir = "Output"
                        if os.path.exists(output_dir):
                            for file in os.listdir(output_dir):
                                if file.endswith('.exe'):
                                    installer_path = os.path.join(output_dir, file)
                                    success_print(f"Installer built: {installer_path}")
                                    return installer_path
                        
                        success_print("Installer build completed (path unknown)")
                        return "installer_built"
                    else:
                        error_print(f"Installer build failed: {result.stderr}")
                        return None
                else:
                    warning_print("Inno Setup not found, skipping installer build")
                    return None
            else:
                warning_print("No installer script found")
                return None
                
        except subprocess.TimeoutExpired:
            error_print("Installer build timed out")
            return None
        except Exception as e:
            error_print(f"Installer build failed: {e}")
            return None
    
    def create_database_config(self) -> str:
        """Create database configuration file."""
        debug_print("Creating database configuration...")
        
        try:
            os.makedirs("config", exist_ok=True)
            
            config = {
                "database": {
                    "type": "postgresql",
                    "host": "your_synology_ip",
                    "port": 5432,
                    "database": "dataviewer",
                    "username": "your_username",
                    "password": "your_password"
                },
                "app_settings": {
                    "auto_update_check": True,
                    "update_check_interval": 24,
                    "enable_telemetry": False
                },
                "build_info": {
                    "build_date": datetime.now().isoformat(),
                    "build_version": "production",
                    "build_mode": "release"
                }
            }
            
            config_path = 'config/database_config.json'
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=2)
            
            debug_print(f"Created database config: {config_path}")
            return config_path
            
        except Exception as e:
            error_print(f"Database config creation failed: {e}")
            return ""


class UpdateService:
    """
    Consolidated service for all update operations including version checking,
    database updates, version management, and release workflows.
    """
    
    def __init__(self, current_version: str = "3.0.0", github_repo: str = "acbecquet/DataViewer",
                 database_service=None, file_service=None):
        """Initialize the update service."""
        debug_print("Initializing UpdateService")
        
        # Basic configuration
        self.current_version = current_version
        self.github_repo = github_repo
        self.github_api_base = "https://api.github.com"
        self.releases_endpoint = f"{self.github_api_base}/repos/{self.github_repo}/releases"
        
        # Service dependencies
        self.database_service = database_service
        self.file_service = file_service
        
        # Update checking configuration
        self.last_check: Optional[datetime] = None
        self.check_interval = timedelta(hours=24)  # Check once per day
        self.debug_mode = True
        
        # Initialize component managers
        self.version_manager = VersionManager()
        self.file_version_updater = FileVersionUpdater()
        self.database_update_manager = DatabaseUpdateManager(database_service, file_service)
        self.github_release_manager = GitHubReleaseManager(github_repo)
        self.build_workflow_manager = BuildWorkflowManager()
        
        # Update tracking
        self.update_history = []
        self.last_update_check_result = None
        
        debug_print("UpdateService initialized successfully")
        debug_print(f"Current version: {current_version}")
        debug_print(f"GitHub repo: {github_repo}")
    
    # ===================== VERSION CHECKING =====================
    
    def check_for_updates(self, force_check: bool = False) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """Check for available updates from GitHub releases."""
        debug_print("Checking for updates...")
        
        try:
            # Rate limiting check
            if not force_check and self.last_check:
                time_since_check = datetime.now() - self.last_check
                if time_since_check < self.check_interval:
                    remaining = self.check_interval - time_since_check
                    return False, None, f"Rate limited. Next check in {remaining}"
            
            # Get latest release info
            release_info = self.github_release_manager.get_latest_release_info()
            self.last_check = datetime.now()
            
            if not release_info:
                return False, None, "Failed to fetch release information"
            
            latest_version = release_info['version']
            
            # Compare versions
            if self.version_manager.is_newer_version(latest_version, self.current_version):
                self.last_update_check_result = release_info
                
                debug_print(f"Update available: {self.current_version} -> {latest_version}")
                return True, release_info, "Update available"
            else:
                debug_print("No updates available")
                return False, None, "No updates available"
            
        except requests.exceptions.RequestException as e:
            error_msg = f"Network error checking for updates: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = f"Failed to check for updates: {e}"
            debug_print(f"ERROR: {error_msg}")
            traceback.print_exc()
            return False, None, error_msg
    
    def get_update_info(self) -> Optional[Dict[str, Any]]:
        """Get information about the last found update."""
        return self.last_update_check_result
    
    # ===================== DATABASE UPDATES =====================
    
    def track_file_modification(self, file_data: Dict[str, Any], modification_type: str = "data_change"):
        """Track that a file has been modified and needs database update."""
        self.database_update_manager.track_file_modification(file_data, modification_type)
    
    def get_modified_files(self) -> List[Dict[str, Any]]:
        """Get list of files that have been modified and need database updates."""
        return self.database_update_manager.get_modified_files()
    
    def update_database_with_modifications(self, progress_callback: Optional[Callable] = None) -> Tuple[bool, str]:
        """
        Update database with all tracked file modifications.
        
        Args:
            progress_callback: Optional callback for progress updates (receives percentage)
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        debug_print("Starting database update with tracked modifications")
        
        try:
            modified_files = self.get_modified_files()
            
            if not modified_files:
                return True, "No files have been modified. Nothing to update."
            
            # Update database
            successful_updates, failed_files = self.database_update_manager.update_database_with_modified_files(
                modified_files, progress_callback
            )
            
            # Generate result message
            if failed_files:
                if successful_updates > 0:
                    message = f"Partial success:\n"
                    message += f"✓ Successfully updated: {successful_updates} files\n"
                    message += f"✗ Failed to update: {len(failed_files)} files\n\n"
                    message += "Failed files:\n" + "\n".join([f"• {name}" for name in failed_files])
                    return False, message
                else:
                    message = f"Failed to update all {len(failed_files)} files:\n\n"
                    message += "\n".join([f"• {name}" for name in failed_files])
                    return False, message
            else:
                message = f"Successfully updated {successful_updates} file(s) in the database."
                # Clear modification flags on complete success
                self.database_update_manager.clear_modification_flags()
                return True, message
            
        except Exception as e:
            error_msg = f"Database update failed: {e}"
            debug_print(f"ERROR: {error_msg}")
            traceback.print_exc()
            return False, error_msg
    
    def clear_modification_flags(self):
        """Clear all modification tracking flags."""
        self.database_update_manager.clear_modification_flags()
    
    # ===================== VERSION MANAGEMENT =====================
    
    def get_current_version(self) -> str:
        """Get current application version."""
        # Try to get from setup.py first, then fallback to stored version
        version_from_setup = self.file_version_updater.get_current_version_from_setup()
        return version_from_setup if version_from_setup else self.current_version
    
    def increment_version(self, increment_type: str = 'patch') -> Optional[str]:
        """Increment version number."""
        current = self.get_current_version()
        return self.version_manager.increment_version(current, increment_type)
    
    def update_version_in_files(self, new_version: str) -> List[str]:
        """Update version strings in all relevant source files."""
        return self.file_version_updater.update_version_in_files(new_version)
    
    def compare_versions(self, version1: str, version2: str) -> int:
        """Compare two version strings."""
        return self.version_manager.compare_versions(version1, version2)
    
    # ===================== RELEASE WORKFLOW =====================
    
    def execute_full_release_workflow(self, increment_type: str = 'patch', 
                                    release_notes: str = "",
                                    create_github_release: bool = True,
                                    progress_callback: Optional[Callable] = None) -> Tuple[bool, str]:
        """
        Execute complete release workflow including version increment,
        file updates, building, and GitHub release creation.
        
        Args:
            increment_type: How to increment version ('major', 'minor', 'patch')
            release_notes: Notes for the release
            create_github_release: Whether to create GitHub release
            progress_callback: Optional callback for progress updates
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        debug_print(f"Starting full release workflow (increment: {increment_type})")
        
        try:
            total_steps = 8
            current_step = 0
            
            def update_progress(step_name: str):
                nonlocal current_step
                current_step += 1
                progress = int((current_step / total_steps) * 100)
                debug_print(f"Step {current_step}/{total_steps}: {step_name}")
                if progress_callback:
                    progress_callback(progress)
            
            # Step 1: Get current version and increment
            update_progress("Getting current version")
            current_version = self.get_current_version()
            new_version = self.increment_version(increment_type)
            
            if not new_version:
                return False, f"Failed to increment version from {current_version}"
            
            debug_print(f"Version increment: {current_version} -> {new_version}")
            
            # Step 2: Clean up build artifacts
            update_progress("Cleaning build artifacts")
            if not self.build_workflow_manager.cleanup_build_artifacts():
                warning_print("Build artifact cleanup failed, continuing anyway")
            
            if not self.build_workflow_manager.cleanup_unwanted_folders():
                warning_print("Unwanted folder cleanup failed, continuing anyway")
            
            # Step 3: Create database config
            update_progress("Creating database config")
            config_path = self.build_workflow_manager.create_database_config()
            if config_path:
                debug_print(f"Created config: {config_path}")
            
            # Step 4: Update version in files
            update_progress("Updating version in files")
            updated_files = self.update_version_in_files(new_version)
            
            if not updated_files:
                return False, "No files were updated with new version"
            
            debug_print(f"Updated version in {len(updated_files)} files")
            
            # Step 5: Build executable
            update_progress("Building executable")
            if not self.build_workflow_manager.build_executable():
                warning_print("Executable build failed, continuing with release")
            
            # Step 6: Build installer
            update_progress("Building installer")
            installer_path = self.build_workflow_manager.build_installer()
            
            if not installer_path:
                warning_print("Installer build failed, continuing without installer")
            
            # Step 7: Commit changes to git (optional)
            update_progress("Committing changes")
            try:
                self._commit_version_changes(new_version, updated_files)
            except Exception as e:
                warning_print(f"Git commit failed: {e}")
            
            # Step 8: Create GitHub release
            update_progress("Creating GitHub release")
            github_success = False
            
            if create_github_release:
                github_success = self.github_release_manager.create_github_release(
                    new_version, installer_path, release_notes
                )
                
                if not github_success:
                    warning_print("GitHub release creation failed")
            
            # Update stored version
            self.current_version = new_version
            
            # Generate success message
            success_msg = f"Release workflow completed successfully!\n\n"
            success_msg += f"✅ Version: {current_version} -> {new_version}\n"
            success_msg += f"✅ Updated files: {len(updated_files)}\n"
            
            if installer_path:
                success_msg += f"✅ Installer: {installer_path}\n"
            
            if github_success:
                success_msg += f"✅ GitHub release created\n"
            
            success_msg += f"\n📦 Next steps:\n"
            success_msg += f"1. Test the new version\n"
            success_msg += f"2. Distribute to users\n"
            success_msg += f"3. Update documentation"
            
            debug_print("Full release workflow completed successfully")
            return True, success_msg
            
        except Exception as e:
            error_msg = f"Release workflow failed: {e}"
            debug_print(f"ERROR: {error_msg}")
            traceback.print_exc()
            return False, error_msg
    
    def create_github_release_only(self, version: str, installer_path: Optional[str] = None,
                                 release_notes: str = "") -> Tuple[bool, str]:
        """Create GitHub release without building."""
        try:
            success = self.github_release_manager.create_github_release(version, installer_path, release_notes)
            
            if success:
                return True, f"GitHub release v{version} created successfully"
            else:
                return False, "GitHub release creation failed"
                
        except Exception as e:
            error_msg = f"GitHub release creation failed: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False, error_msg
    
    # ===================== BUILD OPERATIONS =====================
    
    def build_application(self, clean_first: bool = True) -> Tuple[bool, str]:
        """Build application executable and installer."""
        debug_print("Building application")
        
        try:
            # Clean up if requested
            if clean_first:
                if not self.build_workflow_manager.cleanup_build_artifacts():
                    warning_print("Build cleanup failed, continuing anyway")
            
            # Build executable
            if not self.build_workflow_manager.build_executable():
                return False, "Executable build failed"
            
            # Build installer
            installer_path = self.build_workflow_manager.build_installer()
            
            if installer_path:
                return True, f"Build completed successfully. Installer: {installer_path}"
            else:
                return True, "Build completed successfully (no installer created)"
                
        except Exception as e:
            error_msg = f"Build failed: {e}"
            debug_print(f"ERROR: {error_msg}")
            return False, error_msg
    
    # ===================== PRIVATE HELPER METHODS =====================
    
    def _commit_version_changes(self, version: str, updated_files: List[str]):
        """Commit version changes to git."""
        try:
            debug_print("Committing version changes to git")
            
            # Add updated files
            for file in updated_files:
                subprocess.run(['git', 'add', file], check=True)
            
            # Commit
            commit_message = f"Bump version to {version}"
            subprocess.run(['git', 'commit', '-m', commit_message], check=True)
            
            # Push (optional, might fail if no remote)
            try:
                subprocess.run(['git', 'push', 'origin', 'main'], check=True, timeout=30)
                debug_print("Version changes pushed to remote")
            except:
                debug_print("Could not push to remote (continuing anyway)")
            
            debug_print("Version changes committed successfully")
            
        except Exception as e:
            debug_print(f"WARNING: Git commit failed: {e}")
            raise
    
    # ===================== PUBLIC API METHODS =====================
    
    def get_update_service_status(self) -> Dict[str, Any]:
        """Get comprehensive status of the update service."""
        return {
            'current_version': self.get_current_version(),
            'github_repo': self.github_repo,
            'last_check': self.last_check.isoformat() if self.last_check else None,
            'check_interval_hours': self.check_interval.total_seconds() / 3600,
            'pending_updates': len(self.get_modified_files()),
            'last_update_result': self.last_update_check_result,
            'components': {
                'version_manager': True,
                'file_updater': True,
                'database_manager': self.database_service is not None,
                'github_manager': self.github_release_manager.check_github_cli_available(),
                'build_manager': True
            }
        }
    
    def get_supported_operations(self) -> List[str]:
        """Get list of supported update operations."""
        operations = [
            'check_for_updates', 'track_file_modification', 'update_database_with_modifications',
            'increment_version', 'update_version_in_files', 'build_application'
        ]
        
        if self.github_release_manager.check_github_cli_available():
            operations.extend(['create_github_release_only', 'execute_full_release_workflow'])
        
        return operations
    
    def validate_configuration(self) -> Tuple[bool, List[str]]:
        """Validate update service configuration."""
        issues = []
        
        # Check GitHub repo format
        if not re.match(r'^[\w\-\.]+/[\w\-\.]+$', self.github_repo):
            issues.append(f"Invalid GitHub repo format: {self.github_repo}")
        
        # Check version format
        if not self.version_manager.parse_version(self.current_version):
            issues.append(f"Invalid current version format: {self.current_version}")
        
        # Check required files
        if not os.path.exists('setup.py'):
            issues.append("setup.py not found")
        
        # Check GitHub CLI if needed
        if not self.github_release_manager.check_github_cli_available():
            issues.append("GitHub CLI not available (required for releases)")
        
        return len(issues) == 0, issues