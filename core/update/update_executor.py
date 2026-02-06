"""
Update executor for safely replacing program files.
"""
import sys
import os
import tempfile
import subprocess
from pathlib import Path
from typing import Tuple


class UpdateExecutor:
    """
    Execute update by replacing program files safely.
    
    Uses a VBScript-wrapped batch script to avoid showing console window.
    """
    
    def get_current_exe_path(self) -> Path:
        """
        Get path to current executable.
        
        Returns:
            Path to current exe file
        """
        if getattr(sys, "frozen", False):
            return Path(sys.executable)
        else:
            # Development mode - updates only work in frozen mode
            return Path(__file__).resolve().parent.parent.parent / "dist" / "start.exe"
    
    def _create_updater_scripts(self, current_exe: Path, new_exe_path: Path) -> Tuple[Path, Path]:
        """
        Create batch script and VBScript wrapper to perform the update.
        VBScript wrapper hides the console window.
        
        Args:
            current_exe: Path to current executable
            new_exe_path: Path to new executable
            
        Returns:
            Tuple of (batch_script_path, vbscript_path)
        """
        temp_dir = Path(tempfile.gettempdir())
        pid = os.getpid()
        
        # Create batch script
        batch_script = temp_dir / f"updater_{pid}.bat"
        batch_content = f'''@echo off
setlocal enabledelayedexpansion

set "CURRENT_EXE={current_exe}"
set "NEW_EXE={new_exe_path}"

REM Wait for main process to exit (max 60 seconds)
set /a WAIT_COUNT=0
:WAIT_LOOP
timeout /t 2 /nobreak >nul 2>&1
set /a WAIT_COUNT+=1
if !WAIT_COUNT! geq 30 exit /b 1

REM Check if exe is still locked
del /f /q "!CURRENT_EXE!" >nul 2>&1
if exist "!CURRENT_EXE!" goto WAIT_LOOP

REM Wait additional time for processes to terminate
timeout /t 3 /nobreak >nul 2>&1

REM Verify new exe exists
if not exist "!NEW_EXE!" exit /b 1

REM Move new exe to current directory
move /y "!NEW_EXE!" "!CURRENT_EXE!" >nul 2>&1
if errorlevel 1 exit /b 1

REM Verify move was successful
if not exist "!CURRENT_EXE!" exit /b 1

exit /b 0
'''
        
        with open(batch_script, "w", encoding="utf-8") as f:
            f.write(batch_content)
        
        # Create VBScript wrapper to hide console window
        vbscript = temp_dir / f"updater_{pid}.vbs"
        vbscript_content = f'''Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd.exe /c ""{batch_script}""", 0, False
Set WshShell = Nothing
'''
        
        with open(vbscript, "w", encoding="utf-8") as f:
            f.write(vbscript_content)
        
        return batch_script, vbscript
    
    def execute_update(
        self,
        new_exe_path: Path,
        restart_after_update: bool = True
    ) -> bool:
        """
        Execute update by launching updater process.
        
        Args:
            new_exe_path: Path to the new exe file
            restart_after_update: Whether to restart after update (unused, kept for compatibility)
            
        Returns:
            True if updater was launched successfully
            
        Raises:
            FileNotFoundError: If new exe not found
            OSError: If updater cannot be launched
        """
        if not getattr(sys, "frozen", False):
            raise RuntimeError("Updates are only supported in compiled (frozen) mode")
        
        current_exe = self.get_current_exe_path()
        if not current_exe.exists():
            raise FileNotFoundError(f"Current exe not found: {current_exe}")
        
        if not new_exe_path.exists():
            raise FileNotFoundError(f"New exe not found: {new_exe_path}")
        
        try:
            # Create updater scripts
            batch_script, vbscript = self._create_updater_scripts(current_exe, new_exe_path)
            
            # Launch VBScript wrapper (hides console window)
            subprocess.Popen(
                ["wscript.exe", str(vbscript)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            
            return True
            
        except Exception as e:
            raise OSError(f"Failed to launch updater: {str(e)}") from e
    
    def can_update(self) -> bool:
        """
        Check if update can be performed.
        
        Returns:
            True if update is possible (frozen mode)
        """
        return getattr(sys, "frozen", False)
