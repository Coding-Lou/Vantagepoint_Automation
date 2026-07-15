"""
Worker for Rev-Gen Report Export task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.report_export_adapter import ReportExportAdapter


class ReportExportWorker(BaseWorker):
    """
    Worker for executing Rev-Gen report export task.
    """

    def __init__(self, params: Dict[str, Any]):
        super().__init__("report_export", params, auto_login_enabled=True)
        self.adapter = ReportExportAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit,
        )

    def execute(self) -> Dict[str, Any]:
        if self.is_cancelled:
            return {"success": False, "message": "Task was cancelled before execution"}
        try:
            return self.adapter.execute(**self.params)
        except Exception as e:
            return {"success": False, "message": f"Report export task failed: {str(e)}"}
