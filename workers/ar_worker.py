"""
Worker for AR (Accounts Receivable) task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.ar_adapter import ARAdapter


class ARWorker(BaseWorker):
    """
    Worker for executing AR statement generation task.
    """
    
    def __init__(self, params: Dict[str, Any]):
        """
        Initialize AR worker.
        
        Args:
            params: Task parameters dictionary
        """
        super().__init__("ar.generate_statements", params, auto_login_enabled=True)
        self.adapter = ARAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit
        )
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute AR task.
        
        Returns:
            Result dictionary
        """
        # Check if cancelled before starting
        if self.is_cancelled:
            return {
                "success": False,
                "message": "Task was cancelled before execution"
            }
        
        try:
            #region agent log
            from pathlib import Path
            import json, time
            DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")
            def _agent_log_local(message: str, data: Dict[str, Any]):
                payload = {
                    "sessionId": "debug-session",
                    "runId": "pre-fix",
                    "hypothesisId": "H1",
                    "location": "ARWorker.execute",
                    "message": message,
                    "data": data,
                    "timestamp": int(time.time() * 1000),
                }
                try:
                    DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
                    with DEBUG_LOG_PATH.open("a", encoding="utf-8") as log_file:
                        log_file.write(json.dumps(payload, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            _agent_log_local("ARWorker execute start", {"params_keys": list(self.params.keys())})
            #endregion
            
            # Check again before calling adapter
            if self.is_cancelled:
                return {
                    "success": False,
                    "message": "Task was cancelled"
                }
            
            result = self.adapter.execute(
                statement_date=self.params.get("statement_date"),
                mail_from=self.params.get("mail_from"),
                mail_cc=self.params.get("mail_cc"),
                mail_subject=self.params.get("mail_subject"),
                mail_body=self.params.get("mail_body")
            )
            #region agent log
            _agent_log_local("ARWorker execute result", {"success": bool(result.get("success"))})
            #endregion
            return result
        except Exception as e:
            #region agent log
            _agent_log_local("ARWorker exception", {"error": str(e)})
            #endregion
            return {
                "success": False,
                "message": f"AR task execution failed: {str(e)}"
            }
