"""
Adapter for Project Status task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import json
import time

from domain.adapters.base_adapter import BaseAdapter
import core.project_status as ps_module
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


class ProjectStatusAdapter(BaseAdapter):
    """
    Adapter for project status report generation task.
    
    Wraps the original project_status.py script logic without modifying it.
    """
    
    def execute(
        self,
        period: str,
        use_filter: bool,
        project_names: str,
        start_year: Optional[int] = None,
        download_invoices: bool = True,
        download_project_earnings: bool = True,
        download_expenses: bool = True,
        download_labor_hours: bool = True,
        download_office_earnings: bool = True,
        download_open_po: bool = True
    ) -> Dict[str, Any]:
        """
        Execute project status report generation task.
        
        Args:
            period: Accounting period in YYYYMM format (e.g., 202607)
            use_filter: If True, filter by charge type (Regular) and created date; if False, use manual project names
            project_names: Comma-separated project names string (only used when use_filter is False)
            start_year: Year to use for start date filter (only used when use_filter is True). 
                       If None, defaults to current year - 3.
            download_invoices: Whether to download invoice register
            download_project_earnings: Whether to download project earnings
            download_expenses: Whether to download expenses
            download_labor_hours: Whether to download labor hours
            download_office_earnings: Whether to download office earnings
            download_open_po: Whether to download open purchase orders
            
        Returns:
            Result dictionary with success status and output file path
        """
        try:
            #region agent log
            _agent_log(
                "H1",
                "ProjectStatusAdapter.execute",
                "Project Status adapter execute called",
                {
                    "has_period": bool(period),
                    "use_filter": use_filter,
                    "has_projects": bool(project_names),
                },
            )
            #endregion
            
            try:
                # Log start
                if self.log_callback:
                    self.log_callback("INFO", f"Starting project status report generation for period: {period}")
                
                # CRITICAL: Update HEADERS in ps_module to ensure fresh authentication
                ps_module.HEADERS = util_module.set_headers()
                if self.log_callback:
                    self.log_callback("INFO", "Updated authentication headers")
                   
                # Set period
                if self.log_callback:
                    self.log_callback("INFO", f"Setting accounting period to: {period}")
                util_module.change_period(period)
                
                # Determine filter mode and set searchOptions accordingly
                projects = []
                if use_filter:
                    # Filter by charge type (Regular) and created date
                    from datetime import date
                    # Use provided start_year or default to current year - 3
                    if start_year is None:
                        start_year = date.today().year - 3
                    start_date = f"{start_year}-01-01T00:00:00"
                    ps_module.searchOptions = [
                        {
                            "name": "CreateDate",
                            "value": start_date,
                            "type": "datetime",
                            "seq": 1,
                            "tableName": "PR",
                            "opp": ">",
                            "condition": "and",
                            "searchLevel": 1,
                            "valueDescription": ""
                        },
                        {
                            "name": "ChargeType",
                            "value": "R",
                            "type": "dropdown",
                            "seq": 2,
                            "tableName": "PR",
                            "condition": "and",
                            "searchLevel": 1,
                            "valueDescription": "Regular"
                        }
                    ]
                    if self.log_callback:
                        self.log_callback("INFO", f"Using filter mode: charge type (Regular) and created date after {start_date}")
                else:
                    # Manual project input mode
                    if self.log_callback:
                        self.log_callback("INFO", f"Projects: {project_names}")
                    
                    # Parse project names
                    projects = [p.strip() for p in project_names.split(",") if p.strip()]
                    if not projects:
                        return {
                            "success": False,
                            "message": "No valid project names provided"
                        }
                    
                    if self.log_callback:
                        self.log_callback("INFO", f"Processing {len(projects)} project(s): {', '.join(projects)}")
                    
                    # Assemble projects (this sets the global searchOptions)
                    ps_module.searchOptions = util_module.assamble_projects(projects)
                
                # Ensure project status folder exists
                import os
                ps_dir = "project status"
                util_module.check_folder(ps_dir)
                util_module.clear_folder(ps_dir)
                if self.log_callback:
                    self.log_callback("INFO", f"Prepared output directory: {ps_dir}")
                copy_to_templete = download_expenses and download_invoices and download_labor_hours and download_project_earnings and download_open_po
                # Initialize output
                ps_module.update_templete(r"C:\\temp\\project status", copy_to_templete)
                if self.log_callback:
                    self.log_callback("INFO", "Initialized output workbook")
                
                # Download reports based on user selection
                if download_invoices:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Invoice Register...")
                    ps_module.download_invoices(copy_to_templete)
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Invoice Register (not selected)")
                
                if download_project_earnings:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Project Earnings...")
                    ps_module.download_earnings(copy_to_templete)
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Project Earnings (not selected)")
                
                if download_expenses:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Project Expenses...")
                    ps_module.download_expenses(copy_to_templete)
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Project Expenses (not selected)")
                
                if download_labor_hours:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Labor Hours...")
                    ps_module.download_labor_hours(copy_to_templete)
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Labor Hours (not selected)")

                if download_open_po:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Purchase Orders...")
                    # Only download purchase orders if we have projects (manual mode)
                    # Purchase orders require specific project names, so skip if using filter mode
                    if projects:
                        ps_module.downloand_open_po(projects, copy_to_templete)
                    else:
                        if self.log_callback:
                            self.log_callback("WARNING", "Skipping purchase orders (filter mode - no specific projects)")
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Purchase Orders (not selected)")

                if download_office_earnings:
                    if self.log_callback:
                        self.log_callback("INFO", "Downloading Office Earnings...")
                    ps_module.download_office_earnings()
                else:
                    if self.log_callback:
                        self.log_callback("INFO", "Skipping Office Earnings (not selected)")
                
                if copy_to_templete:
                    ps_module.copy_to_template()
                
                ps_module.move_to_folder()

                import subprocess
                folder_path = os.path.join(os.getcwd(), "project status")
                subprocess.Popen(["explorer", folder_path])

                # Get output file path
                from datetime import date
                output_file = os.path.join(ps_dir, f"output_{date.today().strftime('%Y-%m-%d')}.xlsx")
                
                if self.log_callback:
                    self.log_callback("SUCCESS", f"Project status report generated: {output_file}")
                
                return {
                    "success": True,
                    "output_file": output_file,
                    "message": f"Project status report completed. Output: {output_file}"
                }
                
            except Exception as e:
                #region agent log
                _agent_log(
                    "H1",
                    "ProjectStatusAdapter.execute",
                    "Project Status adapter raised exception",
                    {"error": str(e)},
                )
                #endregion
                if self.log_callback:
                    self.log_callback("ERROR", f"Project Status task failed: {str(e)}")
                return {
                    "success": False,
                    "message": f"Project Status task failed: {str(e)}"
                }
                
        except Exception as e:
            #region agent log
            _agent_log(
                "H1",
                "ProjectStatusAdapter.execute",
                "Project Status adapter outer exception",
                {"error": str(e)},
            )
            #endregion
            if self.log_callback:
                self.log_callback("ERROR", f"Project Status task failed: {str(e)}")
            return {
                "success": False,
                "message": f"Project Status task failed: {str(e)}"
            }
