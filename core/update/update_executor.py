"""
Update executor for safely replacing program files.
"""
import sys
import os
import shutil
import subprocess
import time
import tempfile
from pathlib import Path
from typing import Optional


class UpdateExecutor:
    """
    Execute update by replacing program files safely.
    
    Uses a separate updater process to avoid file locking issues.
    Creates a temporary Python script to perform the update if updater.exe is not available.
    """
    
    def __init__(self, updater_exe_name: str = "updater.exe"):
        """
        Initialize update executor.
        
        Args:
            updater_exe_name: Name of the updater executable (optional, will create script if not found)
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
    
    def _create_updater_script(self, current_exe: Path, new_exe_path: Path) -> Path:
        """
        Create a temporary batch script to perform the update on Windows.
        Batch scripts are more reliable for replacing running executables.
        
        Args:
            current_exe: Path to current executable
            new_exe_path: Path to new executable
            
        Returns:
            Path to created updater script
        """
        if sys.platform == "win32":
            # Use batch script on Windows for better reliability
            exe_name = current_exe.name
            script_content = f'''@echo off
REM Temporary updater batch script to replace executable.
REM This script waits for the current process to exit, deletes the old exe,
REM and moves the new exe from C:\\temp to the current directory.
setlocal enabledelayedexpansion

set "CURRENT_EXE={current_exe}"
set "NEW_EXE={new_exe_path}"
set "EXE_DIR={current_exe.parent}"
set "EXE_NAME={exe_name}"

REM Wait for the main process to exit (max 60 seconds)
REM PyInstaller apps need time to fully release temp files
set /a WAIT_COUNT=0
:WAIT_LOOP
timeout /t 2 /nobreak >nul 2>&1
set /a WAIT_COUNT+=1
if !WAIT_COUNT! geq 30 goto WAIT_FAILED

REM Check if the exe file is still locked by trying to delete it
del /f /q "!CURRENT_EXE!" >nul 2>&1
if exist "!CURRENT_EXE!" (
    REM File is still locked, wait more
    goto WAIT_LOOP
)

REM Wait additional time to ensure all processes are fully terminated
timeout /t 3 /nobreak >nul 2>&1

REM Verify the new exe exists in C:\\temp
if not exist "!NEW_EXE!" (
    echo Error: New executable not found at !NEW_EXE!
    exit /b 1
)

REM Move new exe from C:\\temp to current directory
move /y "!NEW_EXE!" "!CURRENT_EXE!" >nul 2>&1
if errorlevel 1 (
    echo Error: Failed to move new executable to current directory
    exit /b 1
)

REM Verify the move was successful
if not exist "!CURRENT_EXE!" (
    echo Error: Executable not found after move operation
    exit /b 1
)

exit /b 0

:WAIT_FAILED
echo Error: Could not unlock executable file. Update failed.
exit /b 1
'''
            # Create temporary batch file in system temp directory
            temp_dir = Path(tempfile.gettempdir())
            script_path = temp_dir / f"updater_{os.getpid()}.bat"
            
            with open(script_path, "w", encoding="utf-8") as f:
                f.write(script_content)
            
            return script_path
        else:
            # Use Python script on non-Windows platforms
            script_content = f'''"""
Temporary updater script to replace executable and restart application.
"""
import sys
import os
import shutil
import time
import subprocess
from pathlib import Path

def main():
    current_exe = Path(r"{current_exe}")
    new_exe_path = Path(r"{new_exe_path}")
    
    # Wait a bit for the main process to exit
    time.sleep(2)
    
    # Wait for the current exe to be unlocked (max 30 seconds)
    max_wait = 30
    waited = 0
    while waited < max_wait:
        try:
            # Try to rename the file to check if it's locked
            temp_test = current_exe.with_suffix(current_exe.suffix + ".test")
            if temp_test.exists():
                temp_test.unlink()
            current_exe.rename(temp_test)
            temp_test.rename(current_exe)
            break
        except (PermissionError, IOError, OSError):
            time.sleep(0.5)
            waited += 0.5
    
    if waited >= max_wait:
        print("Error: Could not unlock executable file. Update failed.", file=sys.stderr)
        sys.exit(1)
    
    # Replace old exe with new exe
    try:
        temp_exe = current_exe.with_suffix(current_exe.suffix + ".old")
        
        if temp_exe.exists():
            try:
                temp_exe.unlink()
            except Exception:
                pass
        shutil.move(str(current_exe), str(temp_exe))
        shutil.copy2(str(new_exe_path), str(current_exe))
        
        for _ in range(5):
            try:
                temp_exe.unlink()
                break
            except Exception:
                time.sleep(0.5)
        
        try:
            new_exe_path.unlink()
        except Exception:
            pass
        
        subprocess.Popen([str(current_exe)], cwd=str(current_exe.parent))
        
    except Exception as e:
        print(f"Update failed: {{e}}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
'''
            # Create temporary script file
            temp_dir = Path(tempfile.gettempdir())
            script_path = temp_dir / f"updater_{os.getpid()}.py"
            
            with open(script_path, "w", encoding="utf-8") as f:
                f.write(script_content)
            
            return script_path
    
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
            FileNotFoundError: If new exe not found
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
        
        try:
            # Try to use existing updater.exe if available
            updater_path = self.get_updater_path()
            
            if updater_path and updater_path.exists():
                # Use existing updater
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
            else:
                # Create temporary updater script
                updater_script = self._create_updater_script(current_exe, new_exe_path)
                
                # Launch the updater script
                if sys.platform == "win32":
                    # On Windows, use batch script directly
                    creation_flags = subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS
                    subprocess.Popen(
                        ["cmd.exe", "/c", str(updater_script)],
                        creationflags=creation_flags,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                else:
                    # On non-Windows, use Python script
                    subprocess.Popen(
                        [sys.executable, str(updater_script)],
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
        if not getattr(sys, "frozen", False):
            return False
        
        # Update is always possible in frozen mode (we can create updater script)
        return True
