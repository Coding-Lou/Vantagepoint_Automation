"""
Worker for Statement Check task execution.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.statement_check_adapter import StatementCheckAdapter


class StatementCheckWorker(BaseWorker):
    """
    Worker for executing Statement Check task.
    """
    
    def __init__(self, params: Dict[str, Any]):
        """
        Initialize Statement Check worker.
        
        Args:
            params: Task parameters dictionary
        """
        super().__init__("statement_check.check", params, auto_login_enabled=True)
        self.adapter = StatementCheckAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit
        )
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute Statement Check task.
        
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
            # Check again before calling adapter
            if self.is_cancelled:
                return {
                    "success": False,
                    "message": "Task was cancelled"
                }
            
            result = self.adapter.execute(
                vendor_key=self.params.get("vendor_key"),
                invoice_numbers=self.params.get("invoice_numbers")
            )
            
            return result
        except Exception as e:
            return {
                "success": False,
                "message": f"Statement check task execution failed: {str(e)}"
            }
