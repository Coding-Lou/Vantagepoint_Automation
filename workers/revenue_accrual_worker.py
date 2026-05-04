"""
Worker for Revenue Accrual task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.revenue_accrual_adapter import RevenueAccrualAdapter


class RevenueAccrualWorker(BaseWorker):
    """
    Worker for executing revenue accrual generation task.
    """

    def __init__(self, params: Dict[str, Any]):
        """
        Initialize Revenue Accrual worker.

        Args:
            params: Task parameters dictionary
        """
        super().__init__("revenue_accrual.generate_report", params, auto_login_enabled=True)
        self.adapter = RevenueAccrualAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit,
        )

    def execute(self) -> Dict[str, Any]:
        """
        Execute Revenue Accrual task.

        Returns:
            Result dictionary
        """
        # Check if cancelled before starting
        if self.is_cancelled:
            return {
                "success": False,
                "message": "Task was cancelled before execution",
            }

        try:
            # Check again before calling adapter
            if self.is_cancelled:
                return {
                    "success": False,
                    "message": "Task was cancelled",
                }

            result = self.adapter.execute(
                period=self.params.get("period"),
                pkey=self.params.get("pkey", ""),
                option_name=self.params.get("option_name", ""),
            )
            return result
        except Exception as e:
            return {
                "success": False,
                "message": f"Revenue Accrual task execution failed: {str(e)}",
            }

