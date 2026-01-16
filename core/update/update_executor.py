"""
Update executor for safely replacing program files.
"""
import sys
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Optional


class UpdateExecutor:
    """
    Execute update by replacing program files safely.
    
    Uses a separate updater process to avoid file locking issues.
    """
    
    def __init__(self, updater_exe_name: str = "updater.exe"):
        """
        Initialize update executor.
        
        Args:
            updater_exe_name: Name of the updater executable
        """
        self.updater_exe_name = updater_exe_name
    
    def get_current_exe_path(self) -> Path:
        """
        Get path to current executable.
        
        Returns:
            Path to current exe file
        """
        if getattr(sys, "frozen", False):
            # Running as compiled exe
            return Path(sys.executable)
        else:
            # Running as script (development mode)
            # Return a placeholder - updates only work in frozen mode
            return Path(__file__).resolve().parent.parent.parent / "dist" / "start.exe"
    
    def get_updater_path(self) -> Optional[Path]:
        """
        Get path to updater executable.
        
        Returns:
            Path to updater exe or None if not found
        """
        if getattr(sys, "frozen", False):
            # Updater should be in same directory as main exe
            exe_dir = Path(sys.executable).parent
            updater_path = exe_dir / self.updater_exe_name
            if updater_path.exists():
                return updater_path
        else:
            # Development mode: look for updater in tools directory
            tools_dir = Path(__file__).resolve().parent.parent.parent / "tools"
            updater_path = tools_dir / "updater.py"
            if updater_path.exists():
                return updater_path
        
        return None
    
    def execute_update(
        self,
        new_exe_path: Path,
        restart_after_update: bool = True
    ) -> bool:
        """
        Execute update by launching updater process.
        
        This method launches a separate updater process that will:
        1. Wait for current process to exit
        2. Replace old exe with new exe
        3. Restart the application
        
        Args:
            new_exe_path: Path to the new exe file
            restart_after_update: Whether to restart after update
            
        Returns:
            True if updater was launched successfully
            
        Raises:
            FileNotFoundError: If updater or new exe not found
            OSError: If updater cannot be launched
        """
        if not getattr(sys, "frozen", False):
            # Updates only work in frozen (compiled) mode
            raise RuntimeError("Updates are only supported in compiled (frozen) mode")
        
        current_exe = self.get_current_exe_path()
        if not current_exe.exists():
            raise FileNotFoundError(f"Current exe not found: {current_exe}")
        
        if not new_exe_path.exists():
            raise FileNotFoundError(f"New exe not found: {new_exe_path}")
        
        updater_path = self.get_updater_path()
        if not updater_path:
            raise FileNotFoundError(f"Updater not found: {self.updater_exe_name}")
        
        try:
            # Launch updater process
            # Updater arguments: [updater_exe, old_exe_path, new_exe_path]
            if updater_path.suffix == ".py":
                # Development mode: run Python script
                subprocess.Popen([
                    sys.executable,
                    str(updater_path),
                    str(current_exe),
                    str(new_exe_path)
                ])
            else:
                # Production mode: run compiled updater
                subprocess.Popen([
                    str(updater_path),
                    str(current_exe),
                    str(new_exe_path)
                ])
            
            return True
            
        except Exception as e:
            raise OSError(f"Failed to launch updater: {str(e)}") from e
    
    def can_update(self) -> bool:
        """
        Check if update can be performed.
        
        Returns:
            True if update is possible (frozen mode and updater exists)
        """
        if not getattr(sys, "frozen", False):
            return False
        
        updater_path = self.get_updater_path()
        return updater_path is not None and updater_path.exists()
