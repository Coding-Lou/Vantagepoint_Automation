from typing import Dict, Any, Optional
from pathlib import Path

from domain.adapters.base_adapter import BaseAdapter
import core.ar as ar_module
import tools.util as util_module
import tools.config_manager as config_manager


class ARAdapter(BaseAdapter):

    def execute(
        self,
        statement_date: str,
        mail_from: Optional[str] = None,
        mail_cc: Optional[str] = None,
        mail_subject: Optional[str] = None,
        mail_body: Optional[str] = None,
    ) -> Dict[str, Any]:
        try:
            overrides = self._apply_config_overrides(mail_from, mail_cc, mail_subject, mail_body)
            try:
                return self._run(statement_date)
            finally:
                self._restore_config(overrides)
        except Exception as e:
            self._log("ERROR", f"AR task failed: {e}")
            return {"success": False, "message": f"AR task failed: {e}"}

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _run(self, statement_date: str) -> Dict[str, Any]:
        self._log("INFO", f"Starting AR statement generation for date: {statement_date}")

        ar_module.HEADERS     = util_module.set_headers()
        ar_module.MAIL_FROM   = util_module.get_config(["AR", "FROM"])
        ar_module.CC          = util_module.get_config(["AR", "CC"])
        ar_module.SUBJECT     = util_module.get_config(["AR", "SUBJECT"])
        ar_module.BODY        = util_module.get_config(["AR", "BODY"])
        ar_module.OPTIONALMSG = util_module.get_config(["AR", "OPTIONALMSG"])
        ar_module.EXCLUDE     = util_module.get_config(["AR", "EXCLUDE"])

        ar_module.ar_init_with_date(statement_date)
        ar_module.init_output()

        self._log("INFO", "Downloading AR statement CSV...")
        ar_module.ar_download_csv()

        self._log("INFO", "Processing AR statements...")
        ar_module.HEADERS = util_module.set_headers()
        zip_client_names = ar_module.ar_process()
        self._log("INFO", f"Processed {ar_module.RECORDS} records")

        if zip_client_names:
            self._log("INFO", f"Creating zip files for {len(zip_client_names)} clients...")
            for client_name in zip_client_names:
                try:
                    client_id = util_module.get_clientID(client_name)
                    ar_module.ar_zipfile(client_id, client_name)
                except Exception as e:
                    self._log("ERROR", f"Failed to create zip for {client_name}: {e}")

        self._log("INFO", "Saving Excel file...")
        excel_file = util_module.save_excel(
            ar_module.wb,
            ar_module.RECORDS,
            folder_path=str(Path(ar_module.ONEDRIVEDIR) / ar_module.WORKDIR),
        )
        self._log("SUCCESS", f"Saved: {excel_file}")

        return {
            "success": True,
            "output_file": excel_file,
            "records_count": ar_module.RECORDS,
            "message": f"AR task completed. {ar_module.RECORDS} records processed.",
        }

    def _apply_config_overrides(
        self,
        mail_from: Optional[str],
        mail_cc: Optional[str],
        mail_subject: Optional[str],
        mail_body: Optional[str],
    ) -> Dict[str, Any]:
        fields = {
            ("AR", "FROM"): mail_from,
            ("AR", "CC"): mail_cc,
            ("AR", "SUBJECT"): mail_subject,
            ("AR", "BODY"): mail_body,
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
