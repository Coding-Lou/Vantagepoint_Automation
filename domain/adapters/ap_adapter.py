from typing import Dict, Any, Optional, List
from pathlib import Path
import os

from domain.adapters.base_adapter import BaseAdapter
import core.ap as ap_module
import tools.util as util_module
import tools.config_manager as config_manager


class APAdapter(BaseAdapter):

    def execute(
        self,
        remittance_date: str,
        exclude_vendors: Optional[List[str]] = None,
        mail_from: Optional[str] = None,
        mail_cc: Optional[str] = None,
        mail_subject: Optional[str] = None,
        mail_body: Optional[str] = None,
    ) -> Dict[str, Any]:
        try:
            overrides = self._apply_config_overrides(
                exclude_vendors, mail_from, mail_cc, mail_subject, mail_body
            )
            try:
                return self._run(remittance_date)
            finally:
                self._restore_config(overrides)
        except Exception as e:
            self._log("ERROR", f"AP task failed: {e}")
            return {"success": False, "message": f"AP task failed: {e}"}

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _run(self, remittance_date: str) -> Dict[str, Any]:
        self._log("INFO", f"Starting AP remittance for date: {remittance_date}")

        ap_module.HEADERS   = util_module.set_headers()
        ap_module.EXCLUDE   = util_module.get_config(["AP", "EXCLUDE"])
        ap_module.MAIL_FROM = util_module.get_config(["AP", "FROM"])
        ap_module.CC        = util_module.get_config(["AP", "CC"])
        ap_module.SUBJECT   = util_module.get_config(["AP", "SUBJECT"])
        ap_module.BODY      = util_module.get_config(["AP", "BODY"])

        export_dir = os.path.join(ap_module.ONEDRIVEDIR, ap_module.WORKDIR, "ap_export")
        util_module.check_folder(export_dir)

        ap_module.init_output()
        self._set_remittance_date(remittance_date)

        self._log("INFO", "Fetching remittance data...")
        payments_data = ap_module.ap_get_remittance()
        if not payments_data:
            return {"success": False, "message": "No remittance data found"}

        self._log("INFO", f"Found {len(payments_data)} records — processing...")
        ap_module.HEADERS = util_module.set_headers()
        ap_module.ap_process_remittance(payments_data)

        self._log("INFO", "Saving Excel file...")
        excel_file = util_module.save_excel(
            ap_module.wb,
            ap_module.RECORDS,
            folder_path=str(Path(ap_module.ONEDRIVEDIR) / ap_module.WORKDIR),
        )
        self._log("SUCCESS", f"Saved: {excel_file}")

        return {
            "success": True,
            "output_file": excel_file,
            "records_count": ap_module.RECORDS,
            "message": f"AP task completed. {ap_module.RECORDS} records processed.",
        }

    def _set_remittance_date(self, date: str) -> None:
        if not hasattr(ap_module, "ap_setup_time_with_date"):
            raise NotImplementedError(
                "ap.py is missing ap_setup_time_with_date(date: str)"
            )
        ap_module.ap_setup_time_with_date(date)

    def _apply_config_overrides(
        self,
        exclude_vendors: Optional[List[str]],
        mail_from: Optional[str],
        mail_cc: Optional[str],
        mail_subject: Optional[str],
        mail_body: Optional[str],
    ) -> Dict[str, Any]:
        """Write any non-None params to config; return originals for later restore."""
        fields = {
            ("AP", "EXCLUDE"): exclude_vendors,
            ("AP", "FROM"): mail_from,
            ("AP", "CC"): mail_cc,
            ("AP", "SUBJECT"): mail_subject,
            ("AP", "BODY"): mail_body,
        }
        originals: Dict[str, Any] = {}
        for keys, value in fields.items():
            if value is None:
                continue
            current = util_module.get_config(list(keys))
            if value != current:
                originals[keys] = current
                config_manager.update_config_value(list(keys), value)
        return originals

    def _restore_config(self, originals: Dict[str, Any]) -> None:
        for keys, value in originals.items():
            config_manager.update_config_value(list(keys), value)
