"""
Worker for Project Status task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.project_status_adapter import ProjectStatusAdapter


class ProjectStatusWorker(BaseWorker):
    """
    Worker for executing project status report generation task.
    """
    
    def __init__(self, params: Dict[str, Any]):
        """
        Initialize Project Status worker.
        
        Args:
            params: Task parameters dictionary
        """
        super().__init__("project_status.generate_report", params, auto_login_enabled=True)
        self.adapter = ProjectStatusAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit
        )
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute Project Status task.
        
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
                    "location": "ProjectStatusWorker.execute",
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
            _agent_log_local("ProjectStatusWorker execute start", {"params_keys": list(self.params.keys())})
            #endregion
            
            # Check again before calling adapter
            if self.is_cancelled:
                return {
                    "success": False,
                    "message": "Task was cancelled"
                }
            
            result = self.adapter.execute(
                period=self.params.get("period"),
                use_filter=self.params.get("use_filter", False),
                project_names=self.params.get("project_names", ""),
                start_year=self.params.get("start_year")
            )
            #region agent log
            _agent_log_local("ProjectStatusWorker execute result", {"success": bool(result.get("success"))})
            #endregion
            return result
        except Exception as e:
            #region agent log
            _agent_log_local("ProjectStatusWorker exception", {"error": str(e)})
            #endregion
            return {
                "success": False,
                "message": f"Project Status task execution failed: {str(e)}"
            }
