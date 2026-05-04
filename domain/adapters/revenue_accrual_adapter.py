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
        pkey: str = "",
        option_name: str = "",
    ) -> Dict[str, Any]:
        """
        Execute revenue accrual task.

        Args:
            period: Accounting period in YYYYMM format (e.g., 202607)
            pkey: PKey of an existing saved search option; empty means generate new
            option_name: Display name of the saved search option

        Returns:
            Result dictionary with success status and optional output file path
        """
        #region agent log
        _agent_log(
            "H1",
            "RevenueAccrualAdapter.execute",
            "RevenueAccrualAdapter.execute entry",
            {"has_period": bool(period), "period": period, "pkey": pkey},
        )
        #endregion

        if not period:
            return {
                "success": False,
                "message": "Accounting period is required",
            }

        try:
            self._log("INFO", f"Starting revenue accrual generation for period: {period}")

            # Refresh authentication headers in the revenue_accrual module.
            # main() also calls set_headers(), but refreshing here ensures the
            # module-level HEADERS is current before any patched sub-calls run.
            ra_module.HEADERS = util_module.set_headers()
            self._log("INFO", "Authentication headers refreshed.")

            # Monkey-patch builtins.input so main() receives the period without
            # blocking on an interactive prompt.
            original_input = builtins.input

            def _fake_input(prompt: str = "") -> str:
                prompt_str = prompt.strip()
                if prompt_str:
                    self._log("INFO", f"{prompt_str} {period}")
                return period

            builtins.input = _fake_input

            try:
                if pkey:
                    self._log("INFO", f"Using existing project scope: {option_name} ({pkey})")
                    self._run_with_existing_scope(pkey, option_name)
                else:
                    self._log("INFO", "Generating new project scope automatically...")
                    ra_module.main()
            finally:
                builtins.input = original_input

            # Try to infer the primary output workbook path.
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

    def _run_with_existing_scope(self, pkey: str, option_name: str) -> None:
        """
        Run main() with project-list scanning patched out.

        Only the three functions that are purely project_list builders are
        replaced with no-ops (download_invoice_YTD, download_labour_details_YTD,
        and the needDownload=False download_GL call).

        download_vp_revgen is intentionally NOT patched: besides adding to
        project_list it also writes VP Rev_gen_*.csv to C:/temp/revenue_accrual,
        which final_step() reads via Excel automation. Skipping it would cause
        a COM "file not found" error in final_step().

        search_options is patched to return the caller-provided (pkey, option_name)
        instead of creating a new ERP search entry.
        """
        _orig_invoice_ytd = ra_module.download_invoice_YTD
        _orig_labour_ytd = ra_module.download_labour_details_YTD
        _orig_download_gl = ra_module.download_GL
        _orig_search_options = ra_module.search_options

        def _gl_skip_project_list(startPeriod, endPeriod, needDownload, baseRecordSelection, fileName=None):
            # needDownload=False calls are only for building project_list; skip them.
            if not needDownload:
                return None
            return _orig_download_gl(startPeriod, endPeriod, needDownload, baseRecordSelection, fileName)

        ra_module.download_invoice_YTD = lambda: None
        ra_module.download_labour_details_YTD = lambda: None
        ra_module.download_GL = _gl_skip_project_list
        ra_module.search_options = lambda project_list: (pkey, option_name)

        try:
            ra_module.main()
        finally:
            ra_module.download_invoice_YTD = _orig_invoice_ytd
            ra_module.download_labour_details_YTD = _orig_labour_ytd
            ra_module.download_GL = _orig_download_gl
            ra_module.search_options = _orig_search_options
            import subprocess, os
            folder_path = os.path.join(os.getcwd(), "revenue_accrual")
            subprocess.Popen(["explorer", folder_path])
