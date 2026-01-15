"""
Adapter for AR (Accounts Receivable) task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import sys
from io import StringIO
import json
import time

from domain.adapters.base_adapter import BaseAdapter
import core.ar as ar_module
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


class ARAdapter(BaseAdapter):
    """
    Adapter for AR statement generation task.
    
    Wraps the original ar.py script logic without modifying it.
    """
    
    def execute(
        self,
        statement_date: str,
        mail_from: Optional[str] = None,
        mail_cc: Optional[str] = None,
        mail_subject: Optional[str] = None,
        mail_body: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Execute AR statement generation task.
        
        Args:
            statement_date: Statement date in YYYY-MM-DD format
            mail_from: Email sender address
            mail_cc: Email CC addresses
            mail_subject: Email subject
            mail_body: Email body (HTML)
            
        Returns:
            Result dictionary with success status and output file path
        """
        try:
            #region agent log
            _agent_log(
                "H1",
                "ARAdapter.execute",
                "AR adapter execute called",
                {
                    "has_date": bool(statement_date),
                },
            )
            #endregion
            # Save original config values for restoration
            original_from = util_module.get_config(["AR", "FROM"])
            original_cc = util_module.get_config(["AR", "CC"])
            original_subject = util_module.get_config(["AR", "SUBJECT"])
            original_body = util_module.get_config(["AR", "BODY"])
            
            try:
                # Update config if provided
                if mail_from:
                    util_module.set_config(["AR", "FROM"], mail_from)
                if mail_cc:
                    util_module.set_config(["AR", "CC"], mail_cc)
                if mail_subject:
                    util_module.set_config(["AR", "SUBJECT"], mail_subject)
                if mail_body:
                    util_module.set_config(["AR", "BODY"], mail_body)
                
                # Log start
                if self.log_callback:
                    self.log_callback("INFO", f"Starting AR statement generation for date: {statement_date}")
                
                # CRITICAL: Update HEADERS in ar_module to ensure fresh authentication
                # ar.py initializes HEADERS at module load time, which may be stale
                # We need to refresh it before execution to use current login credentials
                ar_module.HEADERS = util_module.set_headers()
                if self.log_callback:
                    self.log_callback("INFO", "Updated authentication headers")
                
                # Ensure ar_export folder exists
                import os
                ar_export_dir = "ar_export"
                if not os.path.exists(ar_export_dir):
                    os.makedirs(ar_export_dir, exist_ok=True)
                    if self.log_callback:
                        self.log_callback("INFO", f"Created directory: {ar_export_dir}")
                
                # Initialize AR module with statement date
                self._set_statement_date(statement_date)
                if self.log_callback:
                    self.log_callback("INFO", f"Set statement date to: {statement_date}")
                
                # Initialize output
                ar_module.init_output()
                if self.log_callback:
                    self.log_callback("INFO", "Initialized output workbook")
                
                # Download CSV
                if self.log_callback:
                    self.log_callback("INFO", "Downloading AR statement CSV...")
                ar_module.ar_download_csv()
                
                # Process statements
                if self.log_callback:
                    self.log_callback("INFO", "Processing AR statements...")
                
                # Refresh HEADERS again before processing (in case processing takes long time)
                ar_module.HEADERS = util_module.set_headers()
                
                zip_client_names = ar_module.ar_process()
                
                if self.log_callback:
                    self.log_callback("INFO", f"Processed statements. Records: {ar_module.RECORDS}")
                
                # Process zip files for clients that need them
                if zip_client_names:
                    if self.log_callback:
                        self.log_callback("INFO", f"Creating zip files for {len(zip_client_names)} clients...")
                    for client_name in zip_client_names:
                        try:
                            client_id = util_module.get_clientID(client_name)
                            ar_module.ar_zipfile(client_id, client_name)
                            if self.log_callback:
                                self.log_callback("INFO", f"Created zip file for {client_name}")
                        except Exception as e:
                            if self.log_callback:
                                self.log_callback("ERROR", f"Failed to create zip for {client_name}: {str(e)}")
                
                # Save Excel
                if self.log_callback:
                    self.log_callback("INFO", "Saving Excel file...")
                excel_file = util_module.save_excel(ar_module.wb, ar_module.RECORDS)
                if self.log_callback:
                    self.log_callback("SUCCESS", f"Excel file saved: {excel_file}")
                
                return {
                    "success": True,
                    "output_file": excel_file,
                    "records_count": ar_module.RECORDS,
                    "message": f"AR task completed. {ar_module.RECORDS} records processed."
                }
                
            finally:
                # Restore original config
                util_module.set_config(["AR", "FROM"], original_from)
                util_module.set_config(["AR", "CC"], original_cc)
                util_module.set_config(["AR", "SUBJECT"], original_subject)
                util_module.set_config(["AR", "BODY"], original_body)
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "ARAdapter.execute",
                "AR adapter raised exception",
                {"error": str(e)},
            )
            #endregion
            if self.log_callback:
                self.log_callback("ERROR", f"AR task failed: {str(e)}")
            return {
                "success": False,
                "message": f"AR task failed: {str(e)}"
            }
    
    def _set_statement_date(self, date: str) -> None:
        """
        Set statement date in ar module.
        
        This sets the STATEMENTDATE global variable and initializes AR module.
        
        Args:
            date: Date string in YYYY-MM-DD format
        """
        # Set the global STATEMENTDATE variable
        ar_module.STATEMENTDATE = date
        
        # Initialize AR module (check folder, clear folder, set user settings)
        # We need to replicate ar_init() logic but without input()
        util_module.check_folder("ar_export")
        util_module.clear_folder("ar_export")
        
        # User Setting
        import requests
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/UserSettings"
        payload = {"FW_SEUserOptions":[{"OptionName":"arReviewPaidStatus","OptionValue":"Unpaid","_transType":"U"}]}
        requests.post(url, headers=ar_module.HEADERS, json=payload)
        
        if self.log_callback:
            self.log_callback("INFO", "AR module initialized with statement date")
