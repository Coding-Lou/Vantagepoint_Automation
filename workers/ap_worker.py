"""
Worker for AP (Accounts Payable) task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.ap_adapter import APAdapter


class APWorker(BaseWorker):
    """
    Worker for executing AP remittance generation task.
    """
    
    def __init__(self, params: Dict[str, Any]):
        """
        Initialize AP worker.
        
        Args:
            params: Task parameters dictionary
        """
        super().__init__("ap.generate_remittance", params)
        self.adapter = APAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit
        )
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute AP task.
        
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
                    "location": "APWorker.execute",
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
            _agent_log_local("APWorker execute start", {"params_keys": list(self.params.keys())})
            #endregion
            
            # Check again before calling adapter
            if self.is_cancelled:
                return {
                    "success": False,
                    "message": "Task was cancelled"
                }
            
            result = self.adapter.execute(
                remittance_date=self.params.get("remittance_date"),
                exclude_vendors=self.params.get("exclude_vendors"),
                mail_from=self.params.get("mail_from"),
                mail_cc=self.params.get("mail_cc"),
                mail_subject=self.params.get("mail_subject"),
                mail_body=self.params.get("mail_body")
            )
            #region agent log
            _agent_log_local("APWorker execute result", {"success": bool(result.get("success"))})
            #endregion
            return result
        except Exception as e:
            #region agent log
            _agent_log_local("APWorker exception", {"error": str(e)})
            #endregion
            return {
                "success": False,
                "message": f"AP task execution failed: {str(e)}"
            }
