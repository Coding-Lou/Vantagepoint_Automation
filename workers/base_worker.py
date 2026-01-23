"""
Base worker class for task execution in QThread.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import sys
import json
import time

from PySide6.QtCore import QThread, Signal

from ui.utils.log_redirector import LogRedirector
import tools.util as util_module
import tools.login as login_module
from ui.services.app_context import get_app_context

#region agent log
DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Optional[Dict[str, Any]] = None,
    run_id: str = "pre-fix",
) -> None:
    payload = {
        "sessionId": "debug-session",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data or {},
        "timestamp": int(time.time() * 1000),
    }
    try:
        DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as log_file:
            log_file.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass
#endregion


class BaseWorker(QThread):
    """
    Base worker class for executing tasks in background thread.
    
    All task workers should inherit from this class and implement
    the execute() method.
    """
    
    # Signals for communication with UI thread
    log_signal = Signal(str, str)  # level, message
    progress_signal = Signal(int, int, str)  # current, total, message
    finished = Signal(dict)  # result dictionary
    
    def __init__(self, task_id: str, params: Dict[str, Any], auto_login_enabled: bool = False):
        """
        Initialize worker.
        
        Args:
            task_id: Task identifier
            params: Task parameters dictionary
            auto_login_enabled: Whether to automatically attempt login if not logged in (default: False)
        """
        super().__init__()
        self.task_id = task_id
        self.params = params
        self.is_cancelled = False
        self.auto_login_enabled = auto_login_enabled
        self.log_redirector: Optional[LogRedirector] = None
        self.old_stdout: Optional[Any] = None
        self.old_stderr: Optional[Any] = None
    
    def run(self) -> None:
        """
        Execute task in background thread.
        
        This method is called automatically when thread starts.
        Do not call this method directly.
        """
        try:
            # Setup log redirection
            self._setup_log_redirect()
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker.run",
                "Worker started",
                {"task_id": self.task_id, "params_keys": list(self.params.keys())},
            )
            #endregion
            
            # Check if cancelled before login check
            if self.is_cancelled:
                self.finished.emit({
                    "success": False,
                    "message": "Task was cancelled"
                })
                return
            
            # Check login status
            if not self._check_login():
                #region agent log
                _agent_log(
                    "H2",
                    "BaseWorker.run",
                    "Login check failed",
                    {"task_id": self.task_id, "auto_login_enabled": self.auto_login_enabled},
                )
                #endregion
                
                # Attempt auto-login if enabled
                if self.auto_login_enabled:
                    self.log_signal.emit("INFO", "Not logged in. Attempting automatic login...")
                    if not self._auto_login():
                        self.finished.emit({
                            "success": False,
                            "message": "Login status is invalid and automatic login failed. Please login manually."
                        })
                        return
                    self.log_signal.emit("SUCCESS", "Automatic login successful. Proceeding with task execution...")
                else:
                    self.finished.emit({
                        "success": False,
                        "message": "Login status is invalid. Please login first."
                    })
                    return
            
            # Check if cancelled before execution
            if self.is_cancelled:
                self.finished.emit({
                    "success": False,
                    "message": "Task was cancelled"
                })
                return
            
            # Execute task (implemented by subclass)
            result = self.execute()
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker.run",
                "Worker finished execution",
                {"task_id": self.task_id, "success": bool(result.get("success"))},
            )
            #endregion
            
            # Emit finished signal
            self.finished.emit(result)
            
        except Exception as e:
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker.run",
                "Worker raised exception",
                {"task_id": self.task_id, "error": str(e)},
            )
            #endregion
            self.finished.emit({
                "success": False,
                "message": f"Task execution exception: {str(e)}"
            })
        finally:
            # Restore log redirection
            self._restore_log_redirect()
    
    def _setup_log_redirect(self) -> None:
        """Setup stdout/stderr redirection to Signal."""
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        self.log_redirector = LogRedirector(self.log_signal.emit)
        sys.stdout = self.log_redirector
        sys.stderr = self.log_redirector
    
    def _restore_log_redirect(self) -> None:
        """Restore original stdout/stderr."""
        if self.old_stdout:
            sys.stdout = self.old_stdout
        if self.old_stderr:
            sys.stderr = self.old_stderr
    
    def _check_login(self) -> bool:
        """
        Check login status.
        
        Returns:
            True if logged in, False otherwise
        """
        try:
            return util_module.check_login()
        except Exception:
            return False
    
    def _auto_login(self) -> bool:
        """
        Attempt automatic login.
        
        Returns:
            True if login successful, False otherwise
        """
        try:
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker._auto_login",
                "Auto-login attempt started",
                {"task_id": self.task_id},
            )
            #endregion
            
            self.log_signal.emit("INFO", "Opening browser for authentication...")
            # Call SSO login function
            login_module.sso_login()
            
            # Verify login was successful
            self.log_signal.emit("INFO", "Verifying login status...")
            user_email = util_module.check_login()
            is_logged_in = bool(user_email)
            
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker._auto_login",
                "Auto-login attempt completed",
                {"task_id": self.task_id, "success": is_logged_in, "user_email": user_email if user_email else None},
            )
            #endregion
            
            # Update AppContext if login successful
            if is_logged_in and user_email:
                try:
                    app_context = get_app_context()
                    user_info = {"EMail": user_email, "email": user_email}
                    app_context.set_login_state(True, user_info)
                    self.log_signal.emit("INFO", f"Login state updated in application context")
                except Exception as e:
                    # Log but don't fail the login process
                    self.log_signal.emit("WARNING", f"Failed to update application context: {str(e)}")
            
            return is_logged_in
        except Exception as e:
            #region agent log
            _agent_log(
                "H2",
                "BaseWorker._auto_login",
                "Auto-login exception",
                {"task_id": self.task_id, "error": str(e)},
            )
            #endregion
            self.log_signal.emit("ERROR", f"Automatic login failed: {str(e)}")
            return False
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute task logic.
        
        This method must be implemented by subclasses.
        
        Returns:
            Result dictionary with 'success' and 'message' fields
        """
        raise NotImplementedError("Subclass must implement execute() method")
    
    def cancel(self) -> None:
        """Request task cancellation."""
        self.is_cancelled = True
        self.log_signal.emit("INFO", "Task cancellation requested")
