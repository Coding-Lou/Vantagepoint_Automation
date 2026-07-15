"""
Version checker for comparing local and remote versions.
"""
from typing import Optional, Dict, Any
import requests
import tools.config_manager as config_manager


class VersionChecker:
    """
    Check version information from GitHub Releases API.
    """
    
    def __init__(self, github_repo: str = "Coding-Lou/Vantagepoint_Automation"):
        """
        Initialize version checker.
        
        Args:
            github_repo: GitHub repository in format "owner/repo"
        """
        self.github_repo = github_repo
        self.api_base_url = f"https://api.github.com/repos/{github_repo}"
    
    def get_local_version(self) -> Optional[str]:
        """
        Get current local version from config.json.
        
        Returns:
            Version string or None if not found
        """
        try:
            version = config_manager.get_config_value(["VERSION"])
            return version if isinstance(version, str) else None
        except Exception:
            return None
    
    def get_latest_release_info(self, timeout: int = 10) -> Optional[Dict[str, Any]]:
        """
        Fetch latest release information from GitHub Releases API.
        
        Args:
            timeout: Request timeout in seconds
            
        Returns:
            Release information dictionary or None if failed
            {
                "version": str,  # Version tag (e.g., "v1.0.0" or "2025-12-18_QNAP")
                "tag_name": str,  # Original tag name
                "download_url": str,  # Download URL for the exe asset
                "asset_name": str,  # Asset file name
                "release_notes": str,  # Release notes/description
                "published_at": str,  # Publication timestamp
            }
        """
        try:
            url = f"{self.api_base_url}/releases/latest"
            response = requests.get(url, timeout=timeout)
            response.raise_for_status()
            data = response.json()
            
            # Extract version from tag_name (remove "v" prefix if present)
            tag_name = data.get("tag_name", "")
            version = tag_name.replace("v_", "").replace("v", "").strip()
            
            # Find the exe asset (prefer start.exe, fallback to first asset)
            assets = data.get("assets", [])
            exe_asset = None
            for asset in assets:
                asset_name = asset.get("name", "")
                if asset_name.endswith(".exe"):
                    if asset_name == "start.exe":
                        exe_asset = asset
                        break
                    elif exe_asset is None:
                        exe_asset = asset
            
            if not exe_asset:
                return None
            
            return {
                "version": version,
                "tag_name": tag_name,
                "download_url": exe_asset.get("browser_download_url"),
                "asset_name": exe_asset.get("name", ""),
                "release_notes": data.get("body", ""),
                "published_at": data.get("published_at", ""),
            }
        except requests.RequestException as e:
            return None
        except Exception:
            return None
    
    def is_update_available(self, timeout: int = 10) -> tuple[bool, Optional[Dict[str, Any]]]:
        """
        Check if an update is available.
        
        Args:
            timeout: Request timeout in seconds
            
        Returns:
            Tuple of (is_available, release_info)
            - is_available: True if update is available
            - release_info: Release information dict or None
        """
        local_version = self.get_local_version()
        if not local_version:
            # If no local version, consider update available
            release_info = self.get_latest_release_info(timeout)
            return release_info is not None, release_info
        
        release_info = self.get_latest_release_info(timeout)
        if not release_info:
            return False, None
        
        latest_version = release_info.get("version", "")
        
        # Simple string comparison (assumes version format is consistent)
        # For semantic versioning, you might want to use version comparison library
        is_available = latest_version != local_version
        
        return is_available, release_info if is_available else None
