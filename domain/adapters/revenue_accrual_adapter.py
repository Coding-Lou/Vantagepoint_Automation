"""
Adapter for Revenue Accrual task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import json
import time
import builtins

from domain.adapters.base_adapter import BaseAdapter
import core.revenue_accrual as ra_module
import tools.util as util_module

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


class RevenueAccrualAdapter(BaseAdapter):
    """
    Adapter for revenue accrual generation task.
    
    Wraps the original core.revenue_accrual script logic without modifying it.
    """

    def execute(
        self,
        period: str,
    ) -> Dict[str, Any]:
        """
        Execute revenue accrual task.

        Args:
            period: Accounting period in YYYYMM format (e.g., 202607)

        Returns:
            Result dictionary with success status and optional output file path
        """
        #region agent log
        _agent_log(
            "H1",
            "RevenueAccrualAdapter.execute",
            "RevenueAccrualAdapter.execute entry",
            {"has_period": bool(period), "period": period},
        )
        #endregion

        if not period:
            return {
                "success": False,
                "message": "Accounting period is required",
            }

        try:
            self._log("INFO", f"Starting revenue accrual generation for period: {period}")

            # Refresh authentication headers in the revenue_accrual module
            # This ensures the script uses the current login session.
            ra_module.HEADERS = util_module.set_headers()
            self._log("INFO", "Updated authentication headers for revenue accrual module")

            # Monkey patch built-in input so that the original script can run
            # without any interactive console prompts.
            original_input = builtins.input

            def _fake_input(prompt: str = "") -> str:
                # Log the prompt once to make behavior visible in the UI
                prompt_str = prompt.strip()
                if prompt_str:
                    self._log("INFO", f"{prompt_str} {period}")
                return period

            builtins.input = _fake_input

            try:
                # Execute original main() function from core.revenue_accrual.
                # All prints will be captured by BaseWorker's log redirection.
                ra_module.main()
            finally:
                # Always restore original input to avoid side effects.
                builtins.input = original_input

            # Try to infer the primary output workbook path.
            # The script generates files under the "revenue accural" folder
            # and uses TEMPLETE3 as the final consolidated workbook.
            try:
                from pathlib import Path as _Path

                output_dir = _Path("revenue accural")
                output_file = output_dir / getattr(ra_module, "TEMPLETE3", "")
                output_path = str(output_file) if output_file.name and output_file.exists() else ""
            except Exception:
                output_path = ""

            message = (
                f"Revenue accrual task completed for period {period}."
                + (f" Output workbook: {output_path}" if output_path else "")
            )
            self._log("SUCCESS", message)

            return {
                "success": True,
                "output_file": output_path or None,
                "message": message,
            }

        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "RevenueAccrualAdapter.execute",
                "RevenueAccrualAdapter.execute exception",
                {"error": str(e)},
            )
            #endregion
            self._log("ERROR", f"Revenue accrual task failed: {str(e)}")
            return {
                "success": False,
                "message": f"Revenue accrual task failed: {str(e)}",
            }

