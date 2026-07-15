"""
Worker for checking login status in QThread.
"""
from typing import Dict, Any
from pathlib import Path
import json
import time

from workers.base_worker import BaseWorker
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


class CheckLoginWorker(BaseWorker):
    """
    Worker for checking login status in background thread.
    """
    
    def __init__(self):
        """Initialize check login worker."""
        super().__init__("util.check_login", {})
    
    def execute(self) -> Dict[str, Any]:
        """
        Check login status.
        
        Returns:
            Result dictionary with login status
        """
        try:
            #region agent log
            _agent_log(
                "H1",
                "CheckLoginWorker.execute",
                "Check login worker started",
                {},
            )
            #endregion
            
            # Call original check_login function (no modification to tools.util)
            is_logged_in = util_module.check_login()
            
            # Note: check_login() prints user email, but we can't capture it
            # without modifying tools.util. We'll just return the status.
            
            return {
                "success": True,
                "is_logged_in": is_logged_in,
                "message": "Logged in" if is_logged_in else "Not logged in"
            }
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "CheckLoginWorker.execute",
                "Check login worker exception",
                {"error": str(e)},
            )
            #endregion
            
            return {
                "success": False,
                "is_logged_in": False,
                "message": f"Error checking login: {str(e)}"
            }
