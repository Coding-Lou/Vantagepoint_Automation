"""
Worker for login execution in QThread.
"""
from typing import Dict, Any
from pathlib import Path
import json
import time

from workers.base_worker import BaseWorker
import tools.login as login_module
import tools.util as util_module

#region agent log
DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Dict[str, Any] = None,
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


class LoginWorker(BaseWorker):
    """
    Worker for executing SSO login in background thread.
    """
    
    def __init__(self, headless: bool = False, log_callback=None):
        """
        Initialize login worker.
        
        Args:
            headless: Whether to run browser in headless mode (default: False)
            log_callback: Callback function for log messages (level, message)
        """
        super().__init__("login.sso_login", {})
        self.headless = headless
        self.log_callback = log_callback
    
    def run(self) -> None:
        """
        Override BaseWorker.run() to skip login check.
        
        LoginWorker should not check login status before executing,
        since the purpose is to perform login.
        """
        try:
            # Setup log redirection
            self._setup_log_redirect()
            #region agent log
            _agent_log(
                "H2",
                "LoginWorker.run",
                "Login worker started (skipping login check)",
                {},
            )
            #endregion
            
            # Execute login (skip login check)
            result = self.execute()
            #region agent log
            _agent_log(
                "H2",
                "LoginWorker.run",
                "Login worker finished",
                {"success": bool(result.get("success"))},
            )
            #endregion
            
            # Emit finished signal
            self.finished.emit(result)
            
        except Exception as e:
            #region agent log
            _agent_log(
                "H2",
                "LoginWorker.run",
                "Login worker exception",
                {"error": str(e)},
            )
            #endregion
            error_msg = f"Login worker error: {str(e)}"
            self.log_signal.emit("ERROR", error_msg)
            if self.log_callback:
                self.log_callback("ERROR", error_msg)
            self.finished.emit({
                "success": False,
                "message": f"Login worker exception: {str(e)}",
                "is_logged_in": False
            })
        finally:
            # Restore log redirection
            self._restore_log_redirect()
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute SSO login.
        
        Returns:
            Result dictionary with success status and user info
        """
        try:
            #region agent log
            _agent_log(
                "H1",
                "LoginWorker.execute",
                "Login worker started",
                {"headless": self.headless},
            )
            #endregion
            
            # Emit log messages via log_signal (inherited from BaseWorker)
            self.log_signal.emit("INFO", "Starting SSO login...")
            self.log_signal.emit("INFO", "Opening browser for authentication...")
            if self.log_callback:
                self.log_callback("INFO", "Starting SSO login...")
                self.log_callback("INFO", "Opening browser for authentication...")
            
            # Call original login function (no modification to tools.login)
            login_module.sso_login()
            
            # Check if login was successful
            self.log_signal.emit("INFO", "Login process completed. Verifying...")
            if self.log_callback:
                self.log_callback("INFO", "Login process completed. Verifying...")
            
            is_logged_in = util_module.check_login()
            
            if is_logged_in:
                # Get user email from config or check_login output
                # check_login() prints the email, but we need to extract it
                # Since we can't modify tools.util, we'll just indicate success
                self.log_signal.emit("SUCCESS", "Login successful!")
                if self.log_callback:
                    self.log_callback("SUCCESS", "Login successful!")
                
                return {
                    "success": True,
                    "message": "Login successful",
                    "is_logged_in": True
                }
            else:
                self.log_signal.emit("ERROR", "Login failed. Please try again.")
                if self.log_callback:
                    self.log_callback("ERROR", "Login failed. Please try again.")
                
                return {
                    "success": False,
                    "message": "Login failed. Please check your credentials.",
                    "is_logged_in": False
                }
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "LoginWorker.execute",
                "Login worker exception",
                {"error": str(e)},
            )
            #endregion
            
            error_msg = f"Login error: {str(e)}"
            self.log_signal.emit("ERROR", error_msg)
            if self.log_callback:
                self.log_callback("ERROR", error_msg)
            
            return {
                "success": False,
                "message": f"Login error: {str(e)}",
                "is_logged_in": False
            }
