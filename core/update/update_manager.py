"""
Update manager that coordinates version checking, downloading, and execution.
"""
from typing import Optional, Callable, Dict, Any
from pathlib import Path
import sys
import tempfile
import tools.config_manager as config_manager

from core.update.version_checker import VersionChecker
from core.update.update_downloader import UpdateDownloader
from core.update.update_executor import UpdateExecutor


class UpdateManager:
    """
    High-level update manager that coordinates all update operations.
    """
    
    def __init__(
        self,
        github_repo: str = "Coding-Lou/Vantagepoint_Automation",
        exe_name: str = "start.exe",
        progress_callback: Optional[Callable[[int, int], None]] = None
    ):
        """
        Initialize update manager.
        
        Args:
            github_repo: GitHub repository in format "owner/repo"
            exe_name: Name of the executable file to update
            progress_callback: Optional callback for download progress (current, total)
        """
        self.github_repo = github_repo
        self.exe_name = exe_name
        self.version_checker = VersionChecker(github_repo)
        self.downloader = UpdateDownloader(progress_callback)
        self.executor = UpdateExecutor()
    
    def check_update(self, timeout: int = 10) -> tuple[bool, Optional[Dict[str, Any]]]:
        """
        Check if an update is available.
        
        Args:
            timeout: Request timeout in seconds
            
        Returns:
            Tuple of (is_available, release_info)
        """
        return self.version_checker.is_update_available(timeout)
    
    def download_update(
        self,
        download_url: str,
        save_to_temp: bool = True
    ) -> Path:
        """
        Download update file.
        
        Args:
            download_url: URL to download from
            save_to_temp: If True, save to C:\\temp directory; otherwise save to exe directory
            
        Returns:
            Path to downloaded file
            
        Raises:
            requests.RequestException: On network errors
            OSError: On file system errors
        """
        if save_to_temp:
            # Save to C:\temp directory (fixed path for Windows)
            if sys.platform == "win32":
                temp_dir = Path("C:\\temp")
            else:
                temp_dir = Path(tempfile.gettempdir())
            # Ensure directory exists
            temp_dir.mkdir(parents=True, exist_ok=True)
            save_path = temp_dir / self.exe_name
        else:
            # Save to same directory as current exe
            if getattr(sys, "frozen", False):
                exe_dir = Path(sys.executable).parent
            else:
                exe_dir = Path(__file__).resolve().parent.parent.parent / "dist"
            save_path = exe_dir / f"{self.exe_name}.new"
        
        # Download file
        self.downloader.download(download_url, save_path)
        
        return save_path
    
    def execute_update(self, new_exe_path: Path) -> bool:
        """
        Execute update by launching updater process.
        
        Args:
            new_exe_path: Path to downloaded new exe file
            
        Returns:
            True if updater was launched successfully
        """
        return self.executor.execute_update(new_exe_path, restart_after_update=True)
    
    def update_version_in_config(self, version: str) -> bool:
        """
        Update version number in config.json.
        
        Args:
            version: New version string
            
        Returns:
            True if update successful
        """
        try:
            return config_manager.update_config_value(["VERSION"], version)
        except Exception:
            return False
    
    def perform_full_update(self, timeout: int = 10) -> Dict[str, Any]:
        """
        Perform complete update process: check, download, and execute.
        
        Args:
            timeout: Request timeout in seconds
            
        Returns:
            Result dictionary with status information:
            {
                "success": bool,
                "message": str,
                "latest_version": str or None,
                "local_version": str or None,
            }
        """
        local_version = self.version_checker.get_local_version()
        
        try:
            # Check for updates
            is_available, release_info = self.check_update(timeout)
            
            if not is_available or not release_info:
                return {
                    "success": True,
                    "message": "Already on latest version",
                    "latest_version": release_info.get("version") if release_info else None,
                    "local_version": local_version,
                }
            
            latest_version = release_info.get("version")
            download_url = release_info.get("download_url")
            
            if not download_url:
                return {
                    "success": False,
                    "message": "No download URL found in release",
                    "latest_version": latest_version,
                    "local_version": local_version,
                }
            
            # Check if update can be executed
            if not self.executor.can_update():
                return {
                    "success": False,
                    "message": "Update not supported in current mode (development mode)",
                    "latest_version": latest_version,
                    "local_version": local_version,
                }
            
            # Download update
            new_exe_path = self.download_update(download_url, save_to_temp=True)
            
            # Update version in config before executing update
            # (Updater will restart, so config should be updated now)
            self.update_version_in_config(latest_version)
            
            # Execute update (this will launch updater and exit)
            self.execute_update(new_exe_path)
            
            return {
                "success": True,
                "message": "Update downloaded and updater launched. Application will restart.",
                "latest_version": latest_version,
                "local_version": local_version,
            }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Update failed: {str(e)}",
                "latest_version": None,
                "local_version": local_version,
            }
