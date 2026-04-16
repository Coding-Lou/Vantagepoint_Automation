"""
Adapter for Rev-Gen Report Export task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import os
from datetime import date as _date

from domain.adapters.base_adapter import BaseAdapter
import core.report_export as re_module
import tools.util as util_module


def _make_returning_download_project_list():
    """
    core/report_export.py defines download_project_list twice; Python keeps only
    the second definition (line 504), which does NOT return the CSV path.
    download_contract (line 724) calls download_project_list and expects that
    return value — without it, excel_full_copy receives None and crashes.

    This wrapper calls the real function and then returns the path it would
    have produced, so download_contract works correctly.
    """
    _real = re_module.download_project_list

    def _wrapper(base_folder, need_trim=False):
        _real(base_folder, need_trim)
        return os.path.join(
            str(base_folder),
            f"Project List_{_date.today().strftime('%Y-%m-%d')}.csv",
        )

    return _wrapper

BASE_FOLDER = Path("C:/temp/revenue_model_export")


class ReportExportAdapter(BaseAdapter):
    """
    Adapter for Rev-Gen Report Export task.

    Wraps core/report_export.py functions without modifying them.
    Refreshes HEADERS before every execution to avoid stale auth.
    """

    def execute(
        self,
        start_period: str = "",
        end_period: str = "",
        pkey: str = "",
        option_name: str = "",
        download_project_list: bool = False,
        download_estimate: bool = False,
        download_contract: bool = False,
        download_cost_gl: bool = False,
        download_cost_from_system: bool = False,
        download_billing: bool = False,
        download_billed: bool = False,
        **kwargs,
    ) -> Dict[str, Any]:
        """
        Execute the selected Rev-Gen report downloads.

        Args:
            start_period: Period string e.g. "202601" (used for GL range)
            end_period:   Period string e.g. "202607" (primary period)
            pkey:         PKey of the saved search option
            option_name:  Display name of the saved search option
            download_*:   Boolean flags for each report type
        """
        try:
            # Set ERP active period first
            if end_period:
                util_module.change_period(end_period)
                self._log("INFO", f"Active period set to {end_period}.")

            # Refresh auth — module-level HEADERS is stale after import
            re_module.HEADERS = util_module.set_headers()
            self._log("INFO", "Authentication headers refreshed.")

            # Ensure output folder exists
            BASE_FOLDER.mkdir(parents=True, exist_ok=True)
            base_folder = str(BASE_FOLDER)
            self._log("INFO", f"Output folder: {base_folder}")

            if download_project_list:
                self._log("INFO", "Downloading Project List...")
                re_module.download_project_list(base_folder)
                self._log("SUCCESS", "Project List downloaded.")

            if download_estimate:
                self._log("INFO", "Downloading Estimate...")
                re_module.download_estimate(base_folder)
                self._log("SUCCESS", "Estimate downloaded.")

            if download_contract:
                self._log("INFO", "Downloading Contract...")
                # Patch download_project_list to return the CSV path for the
                # duration of this call (see _make_returning_download_project_list).
                _orig = re_module.download_project_list
                re_module.download_project_list = _make_returning_download_project_list()
                try:
                    re_module.download_contract(base_folder, end_period, pkey, option_name)
                finally:
                    re_module.download_project_list = _orig
                self._log("SUCCESS", "Contract downloaded.")

            if download_cost_gl:
                self._log("INFO", "Downloading Cost GL...")
                re_module.download_cost_GL(base_folder, end_period)
                self._log("SUCCESS", "Cost GL downloaded.")

            if download_cost_from_system:
                self._log("INFO", "Downloading Labor Hours (Cost From System)...")
                re_module.download_project_labor_hrs(base_folder, pkey, option_name)
                self._log("INFO", "Downloading Project Expense (Cost From System)...")
                re_module.download_project_expense(base_folder, pkey, option_name)
                self._log("SUCCESS", "Cost From System downloaded.")

            if download_billing:
                self._log("INFO", "Downloading Billing/Spent/WO/Prop...")
                re_module.download_JTD_Billing(base_folder, pkey, option_name)
                self._log("SUCCESS", "Billing/Spent/WO/Prop downloaded.")

            if download_billed:
                self._log("INFO", "Downloading Billed...")
                re_module.download_billed(base_folder, pkey, option_name, end_period)
                self._log("SUCCESS", "Billed downloaded.")

            self._log("INFO", "Moving files to revenue_accrual folder...")
            re_module.final_step(base_folder)
            self._log("SUCCESS", "Files moved to revenue_accrual folder.")

            return {
                "success": True,
                "message": f"All selected reports exported to {base_folder}",
            }

        except Exception as e:
            self._log("ERROR", f"Report export failed: {str(e)}")
            return {"success": False, "message": f"Report export failed: {str(e)}"}
