"""
Adapter for Statement Check task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import json
import time

from domain.adapters.base_adapter import BaseAdapter
import core.statement_check as statement_check_module

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


class StatementCheckAdapter(BaseAdapter):
    """
    Adapter for Statement Check task.
    
    Wraps the original statement_check.py script logic without modifying it.
    """
    
    def execute(
        self,
        vendor_key: str,
        invoice_numbers: str
    ) -> Dict[str, Any]:
        """
        Execute Statement Check task.
        
        Args:
            vendor_key: Vendor key to check
            invoice_numbers: Comma-separated invoice numbers
            
        Returns:
            Result dictionary with success status
        """
        try:
            #region agent log
            _agent_log(
                "H1",
                "StatementCheckAdapter.execute",
                "Statement check adapter execute called",
                {
                    "vendor_key": vendor_key,
                    "invoice_count": len(invoice_numbers.split(',')) if invoice_numbers else 0,
                },
            )
            #endregion
            
            self._log("INFO", f"Starting statement check for vendor: {vendor_key}")
            self._log("INFO", f"Invoice numbers: {invoice_numbers}")
            
            # Call the main function with parameters
            statement_check_module.main(
                vendor_key=vendor_key,
                invoice_numbers=invoice_numbers
            )
            
            self._log("SUCCESS", "Statement check completed successfully")
            
            return {
                "success": True,
                "message": "Statement check completed successfully"
            }
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "StatementCheckAdapter.execute",
                "Statement check adapter exception",
                {"error": str(e)},
            )
            #endregion
            
            error_msg = f"Statement check failed: {str(e)}"
            self._log("ERROR", error_msg)
            
            return {
                "success": False,
                "message": error_msg
            }
