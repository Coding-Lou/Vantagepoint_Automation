"""
Adapter for AP (Accounts Payable) task.
"""
from typing import Dict, Any, Optional, List
from pathlib import Path
import json
import time

from domain.adapters.base_adapter import BaseAdapter
import core.ap as ap_module
import tools.util as util_module
import tools.config_manager as config_manager

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


class APAdapter(BaseAdapter):
    """
    Adapter for AP remittance generation task.
    
    Wraps the original ap.py script logic without modifying it.
    """
    
    def execute(
        self,
        remittance_date: str,
        exclude_vendors: Optional[List[str]] = None,
        mail_from: Optional[str] = None,
        mail_cc: Optional[str] = None,
        mail_subject: Optional[str] = None,
        mail_body: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Execute AP remittance generation task.
        
        Args:
            remittance_date: Remittance date in YYYY-MM-DD format
            exclude_vendors: List of vendor IDs to exclude
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
                "APAdapter.execute",
                "AP adapter execute called",
                {
                    "has_date": bool(remittance_date),
                    "exclude_count": len(exclude_vendors or []),
                },
            )
            #endregion
            # Save original config values for restoration (only if we need to update)
            original_exclude = None
            original_from = None
            original_cc = None
            original_subject = None
            original_body = None
            
            try:
                # Update config only if provided and different from current config
                # This avoids unnecessary config updates and errors
                if exclude_vendors is not None:
                    original_exclude = util_module.get_config(["AP", "EXCLUDE"])
                    if exclude_vendors != original_exclude:
                        config_manager.update_config_value(["AP", "EXCLUDE"], exclude_vendors)
                if mail_from:
                    original_from = util_module.get_config(["AP", "FROM"])
                    if mail_from != original_from:
                        config_manager.update_config_value(["AP", "FROM"], mail_from)
                if mail_cc:
                    original_cc = util_module.get_config(["AP", "CC"])
                    if mail_cc != original_cc:
                        config_manager.update_config_value(["AP", "CC"], mail_cc)
                if mail_subject:
                    original_subject = util_module.get_config(["AP", "SUBJECT"])
                    if mail_subject != original_subject:
                        config_manager.update_config_value(["AP", "SUBJECT"], mail_subject)
                if mail_body:
                    original_body = util_module.get_config(["AP", "BODY"])
                    if mail_body != original_body:
                        config_manager.update_config_value(["AP", "BODY"], mail_body)
                
                # Log start
                if self.log_callback:
                    self.log_callback("INFO", f"Starting AP remittance generation for date: {remittance_date}")
                
                # CRITICAL: Update HEADERS in ap_module to ensure fresh authentication
                # ap.py initializes HEADERS at module load time, which may be stale
                # We need to refresh it before execution to use current login credentials
                ap_module.HEADERS = util_module.set_headers()
                if self.log_callback:
                    self.log_callback("INFO", "Updated authentication headers")
                
                # Ensure ap_export folder exists
                import os
                ap_export_dir = "ap_export"
                if not os.path.exists(ap_export_dir):
                    os.makedirs(ap_export_dir, exist_ok=True)
                    if self.log_callback:
                        self.log_callback("INFO", f"Created directory: {ap_export_dir}")
                
                # Initialize output
                ap_module.init_output()
                if self.log_callback:
                    self.log_callback("INFO", "Initialized output workbook")
                
                # Set remittance date
                self._set_remittance_date(remittance_date)
                if self.log_callback:
                    self.log_callback("INFO", f"Set remittance date to: {remittance_date}")
                
                # Get remittance data
                if self.log_callback:
                    self.log_callback("INFO", "Fetching remittance data from API...")
                payments_data = ap_module.ap_get_remittance()
                
                if not payments_data:
                    if self.log_callback:
                        self.log_callback("WARNING", "No remittance data found")
                    return {
                        "success": False,
                        "message": "No remittance data found"
                    }
                
                if self.log_callback:
                    self.log_callback("INFO", f"Found {len(payments_data)} remittance records")
                
                # Process remittance
                if self.log_callback:
                    self.log_callback("INFO", "Processing remittance records...")
                    self.log_callback("INFO", "Note: If you see 'Error in download remittence', it may be due to:")
                    self.log_callback("INFO", "  - Authentication token expired (try re-login)")
                    self.log_callback("INFO", "  - Network connectivity issues")
                    self.log_callback("INFO", "  - Invalid payment data format")
                    self.log_callback("INFO", "  - Missing ap_export folder (should be auto-created)")
                
                # Refresh HEADERS again before processing (in case processing takes long time)
                ap_module.HEADERS = util_module.set_headers()
                
                ap_module.ap_process_remittance(payments_data)
                
                if self.log_callback:
                    self.log_callback("INFO", f"Processed {ap_module.RECORDS - 1} records")
                
                # Save Excel
                if self.log_callback:
                    self.log_callback("INFO", "Saving Excel file...")
                excel_file = util_module.save_excel(ap_module.wb, ap_module.RECORDS)
                if self.log_callback:
                    self.log_callback("SUCCESS", f"Excel file saved: {excel_file}")
                
                return {
                    "success": True,
                    "output_file": excel_file,
                    "records_count": ap_module.RECORDS,
                    "message": f"AP task completed. {ap_module.RECORDS} records processed."
                }
                
            finally:
                # Restore original config only if we updated it
                if original_exclude is not None:
                    config_manager.update_config_value(["AP", "EXCLUDE"], original_exclude)
                if original_from is not None:
                    config_manager.update_config_value(["AP", "FROM"], original_from)
                if original_cc is not None:
                    config_manager.update_config_value(["AP", "CC"], original_cc)
                if original_subject is not None:
                    config_manager.update_config_value(["AP", "SUBJECT"], original_subject)
                if original_body is not None:
                    config_manager.update_config_value(["AP", "BODY"], original_body)
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "APAdapter.execute",
                "AP adapter raised exception",
                {"error": str(e)},
            )
            #endregion
            if self.log_callback:
                self.log_callback("ERROR", f"AP task failed: {str(e)}")
            return {
                "success": False,
                "message": f"AP task failed: {str(e)}"
            }
    
    def _set_remittance_date(self, date: str) -> None:
        """
        Set remittance date in ap module.
        
        This requires a small modification to ap.py to support
        parameter passing. For now, we use a workaround.
        
        Args:
            date: Date string in YYYY-MM-DD format
        """
        # Workaround: Set date through module attribute if available
        # Otherwise, we need to modify ap.py to add a function:
        # def ap_setup_time_with_date(date: str):
        #     url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/UserSettings"
        #     payload = {...}  # Same as ap_setup_time but with date parameter
        #     ...
        
        # For minimal modification, we can add this function to ap.py:
        # This is the only modification needed to ap.py
        if hasattr(ap_module, 'ap_setup_time_with_date'):
            #region agent log
            _agent_log(
                "H1",
                "APAdapter._set_remittance_date",
                "Using ap_setup_time_with_date",
                {"date": date},
            )
            #endregion
            ap_module.ap_setup_time_with_date(date)
        else:
            # Fallback: Use original function (requires user input)
            # This won't work in UI context, so we need the modification
            #region agent log
            _agent_log(
                "H1",
                "APAdapter._set_remittance_date",
                "ap_setup_time_with_date missing",
                {"date": date},
            )
            #endregion
            raise NotImplementedError(
                "ap.py needs to be modified to support date parameter. "
                "Add function: ap_setup_time_with_date(date: str)"
            )
