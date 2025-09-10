# services/update_service.py
"""
services/update_service.py
Update checking service.
This will contain the update logic from update_checker.py.
"""

import requests
import json
from typing import Optional, Dict, Any, Tuple
from datetime import datetime, timedelta


class UpdateService:
    """Service for checking and managing application updates."""
    
    def __init__(self, current_version: str = "3.0.0", github_repo: str = "acbecquet/DataViewer"):
        """Initialize the update service."""
        self.current_version = current_version
        self.github_repo = github_repo
        self.github_api_base = "https://api.github.com"
        self.releases_endpoint = f"{self.github_api_base}/repos/{self.github_repo}/releases"
        self.last_check: Optional[datetime] = None
        self.check_interval = timedelta(hours=24)  # Check once per day
        
        print("DEBUG: UpdateService initialized")
        print(f"DEBUG: Current version: {current_version}")
        print(f"DEBUG: GitHub repo: {github_repo}")
    
    def check_for_updates(self) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """Check for available updates."""
        print("DEBUG: UpdateService checking for updates")
        
        try:
            # Check if we need to check (rate limiting)
            if self.last_check and datetime.now() - self.last_check < self.check_interval:
                return False, None, "Update check rate limited"
            
            # Get latest release from GitHub API
            latest_url = f"{self.releases_endpoint}/latest"
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': f'DataViewer-UpdateService/{self.current_version}'
            }
            
            response = requests.get(latest_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            release_data = response.json()
            latest_version = release_data.get('tag_name', '').lstrip('v')
            
            self.last_check = datetime.now()
            
            # Compare versions
            if self._is_newer_version(latest_version, self.current_version):
                update_info = {
                    'version': latest_version,
                    'release_url': release_data.get('html_url'),
                    'download_url': self._get_download_url(release_data),
                    'release_notes': release_data.get('body', ''),
                    'published_at': release_data.get('published_at'),
                    'prerelease': release_data.get('prerelease', False)
                }
                
                print(f"DEBUG: UpdateService found new version: {latest_version}")
                return True, update_info, "Update available"
            else:
                print("DEBUG: UpdateService - no updates available")
                return False, None, "No updates available"
            
        except requests.exceptions.RequestException as e:
            error_msg = f"Network error checking for updates: {e}"
            print(f"ERROR: UpdateService - {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = f"Failed to check for updates: {e}"
            print(f"ERROR: UpdateService - {error_msg}")
            return False, None, error_msg
    
    def _is_newer_version(self, version1: str, version2: str) -> bool:
        """Compare two version strings."""
        try:
            # Simple version comparison (assumes semantic versioning)
            v1_parts = [int(x) for x in version1.split('.')]
            v2_parts = [int(x) for x in version2.split('.')]
            
            # Pad with zeros if needed
            max_len = max(len(v1_parts), len(v2_parts))
            v1_parts.extend([0] * (max_len - len(v1_parts)))
            v2_parts.extend([0] * (max_len - len(v2_parts)))
            
            return v1_parts > v2_parts
            
        except ValueError:
            # Fall back to string comparison if version parsing fails
            return version1 > version2
    
    def _get_download_url(self, release_data: Dict[str, Any]) -> Optional[str]:
        """Extract download URL from release data."""
        try:
            assets = release_data.get('assets', [])
            
            # Look for installer file
            for asset in assets:
                name = asset.get('name', '').lower()
                if any(ext in name for ext in ['.exe', '.msi', '.dmg', '.pkg']):
                    return asset.get('browser_download_url')
            
            # Fallback to release page
            return release_data.get('html_url')
            
        except Exception as e:
            print(f"ERROR: UpdateService failed to get download URL: {e}")
            return None
    
    def download_update(self, download_url: str, output_path: str) -> Tuple[bool, str]:
        """Download an update file."""
        print(f"DEBUG: UpdateService downloading update: {download_url}")
        
        try:
            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        # Progress reporting could go here
                        if total_size > 0:
                            progress = (downloaded / total_size) * 100
                            if downloaded % (1024 * 1024) == 0:  # Log every MB
                                print(f"DEBUG: Download progress: {progress:.1f}%")
            
            print(f"DEBUG: UpdateService downloaded update to: {output_path}")
            return True, "Download completed"
            
        except Exception as e:
            error_msg = f"Failed to download update: {e}"
            print(f"ERROR: UpdateService - {error_msg}")
            return False, error_msg
    
    def get_release_history(self, limit: int = 10) -> Tuple[bool, List[Dict[str, Any]], str]:
        """Get release history from GitHub."""
        print(f"DEBUG: UpdateService getting release history (limit: {limit})")
        
        try:
            url = f"{self.releases_endpoint}?per_page={limit}"
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': f'DataViewer-UpdateService/{self.current_version}'
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            releases = response.json()
            
            release_history = []
            for release in releases:
                release_history.append({
                    'version': release.get('tag_name', '').lstrip('v'),
                    'name': release.get('name'),
                    'published_at': release.get('published_at'),
                    'prerelease': release.get('prerelease', False),
                    'draft': release.get('draft', False),
                    'release_notes': release.get('body', ''),
                    'download_count': sum(asset.get('download_count', 0) for asset in release.get('assets', []))
                })
            
            print(f"DEBUG: UpdateService retrieved {len(release_history)} releases")
            return True, release_history, "Success"
            
        except Exception as e:
            error_msg = f"Failed to get release history: {e}"
            print(f"ERROR: UpdateService - {error_msg}")
            return False, [], error_msg