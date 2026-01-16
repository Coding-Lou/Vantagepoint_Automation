"""
Worker for checking and performing updates in background thread.
"""
from typing import Dict, Any, Optional
from PySide6.QtCore import QThread, Signal

from core.update.update_manager import UpdateManager


class UpdateWorker(QThread):
    """
    Worker thread for checking and performing updates.
    
    Provides signals for UI integration:
    - log_signal: Progress and status messages
    - progress_signal: Download progress
    - finished: Update completion result
    """
    
    log_signal = Signal(str, str)  # level, message
    progress_signal = Signal(int, int)  # current_bytes, total_bytes
    finished = Signal(dict)  # result dictionary
    
    def __init__(
        self,
        github_repo: str = "Coding-Lou/Vantagepoint_Automation",
        exe_name: str = "start.exe"
    ):
        """
        Initialize update worker.
        
        Args:
            github_repo: GitHub repository in format "owner/repo"
            exe_name: Name of the executable file to update
        """
        super().__init__()
        self.github_repo = github_repo
        self.exe_name = exe_name
        self.update_manager: Optional[UpdateManager] = None
        self.is_cancelled = False
    
    def _on_progress(self, current: int, total: int) -> None:
        """Handle download progress callback."""
        self.progress_signal.emit(current, total)
    
    def run(self) -> None:
        """Execute update check and process in background thread."""
        try:
            # Create update manager with progress callback
            self.update_manager = UpdateManager(
                github_repo=self.github_repo,
                exe_name=self.exe_name,
                progress_callback=self._on_progress
            )
            
            # Check for updates
            self.log_signal.emit("INFO", "Checking for updates...")
            is_available, release_info = self.update_manager.check_update(timeout=10)
            
            if self.is_cancelled:
                self.finished.emit({
                    "success": False,
                    "message": "Update check cancelled",
                    "cancelled": True
                })
                return
            
            if not is_available or not release_info:
                local_version = self.update_manager.version_checker.get_local_version()
                self.log_signal.emit("INFO", f"Already on latest version: {local_version}")
                self.finished.emit({
                    "success": True,
                    "message": "Already on latest version",
                    "update_available": False,
                    "local_version": local_version,
                })
                return
            
            latest_version = release_info.get("version")
            local_version = self.update_manager.version_checker.get_local_version()
            
            self.log_signal.emit("INFO", f"Update available: {latest_version} (current: {local_version})")
            self.log_signal.emit("INFO", "Starting download...")
            
            # Download update
            download_url = release_info.get("download_url")
            if not download_url:
                self.log_signal.emit("ERROR", "No download URL found in release")
                self.finished.emit({
                    "success": False,
                    "message": "No download URL found in release",
                    "update_available": True,
                    "latest_version": latest_version,
                    "local_version": local_version,
                })
                return
            
            new_exe_path = self.update_manager.download_update(download_url, save_to_temp=True)
            
            if self.is_cancelled:
                # Clean up downloaded file
                try:
                    new_exe_path.unlink()
                except Exception:
                    pass
                self.finished.emit({
                    "success": False,
                    "message": "Update cancelled",
                    "cancelled": True
                })
                return
            
            self.log_signal.emit("SUCCESS", "Download completed")
            self.log_signal.emit("INFO", "Preparing to install update...")
            
            # Update version in config
            self.update_manager.update_version_in_config(latest_version)
            
            # Execute update (this will launch updater and the app will exit)
            self.update_manager.execute_update(new_exe_path)
            
            self.log_signal.emit("SUCCESS", "Update completed. Application will restart.")
            self.finished.emit({
                "success": True,
                "message": "Update completed. Application will restart.",
                "update_available": True,
                "latest_version": latest_version,
                "local_version": local_version,
            })
            
        except Exception as e:
            error_msg = str(e)
            self.log_signal.emit("ERROR", f"Update failed: {error_msg}")
            self.finished.emit({
                "success": False,
                "message": error_msg,
                "update_available": None,
            })
    
    def cancel(self) -> None:
        """Request update cancellation."""
        self.is_cancelled = True
        self.log_signal.emit("INFO", "Update cancellation requested")


class UpdateCheckWorker(QThread):
    """
    Worker thread for checking updates only (no download/install).
    
    Useful for background update checks without user interaction.
    """
    
    log_signal = Signal(str, str)  # level, message
    finished = Signal(dict)  # result dictionary
    
    def __init__(
        self,
        github_repo: str = "Coding-Lou/Vantagepoint_Automation"
    ):
        """
        Initialize update check worker.
        
        Args:
            github_repo: GitHub repository in format "owner/repo"
        """
        super().__init__()
        self.github_repo = github_repo
        self.update_manager: Optional[UpdateManager] = None
    
    def run(self) -> None:
        """Check for updates in background thread."""
        try:
            self.update_manager = UpdateManager(github_repo=self.github_repo)
            
            self.log_signal.emit("INFO", "Checking for updates...")
            is_available, release_info = self.update_manager.check_update(timeout=10)
            
            local_version = self.update_manager.version_checker.get_local_version()
            
            if is_available and release_info:
                latest_version = release_info.get("version")
                self.log_signal.emit("INFO", f"Update available: {latest_version}")
                self.finished.emit({
                    "success": True,
                    "update_available": True,
                    "latest_version": latest_version,
                    "local_version": local_version,
                    "release_info": release_info,
                })
            else:
                self.log_signal.emit("INFO", f"Already on latest version: {local_version}")
                self.finished.emit({
                    "success": True,
                    "update_available": False,
                    "latest_version": local_version,
                    "local_version": local_version,
                })
                
        except Exception as e:
            error_msg = str(e)
            self.log_signal.emit("ERROR", f"Update check failed: {error_msg}")
            self.finished.emit({
                "success": False,
                "message": error_msg,
                "update_available": None,
            })
