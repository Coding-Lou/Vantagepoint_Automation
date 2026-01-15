"""
Adapter for Project Status task.
"""
from typing import Dict, Any, Optional
from pathlib import Path
import sys
from io import StringIO
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
        project_names: str
    ) -> Dict[str, Any]:
        """
        Execute project status report generation task.
        
        Args:
            period: Accounting period in YYYYMM format (e.g., 202607)
            project_names: Comma-separated project names string
            
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
                    "has_projects": bool(project_names),
                },
            )
            #endregion
            
            try:
                # Log start
                if self.log_callback:
                    self.log_callback("INFO", f"Starting project status report generation for period: {period}")
                    self.log_callback("INFO", f"Projects: {project_names}")
                
                # CRITICAL: Update HEADERS in ps_module to ensure fresh authentication
                ps_module.HEADERS = util_module.set_headers()
                if self.log_callback:
                    self.log_callback("INFO", "Updated authentication headers")
                
                # Print available periods (for reference, but don't require user input)
                if self.log_callback:
                    self.log_callback("INFO", "Fetching available accounting periods...")
                ps_module.print_period()
                
                # Set period
                if self.log_callback:
                    self.log_callback("INFO", f"Setting accounting period to: {period}")
                util_module.change_period(period)
                
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
                
                # Initialize output
                ps_module.init_output()
                if self.log_callback:
                    self.log_callback("INFO", "Initialized output workbook")
                
                # Download reports
                if self.log_callback:
                    self.log_callback("INFO", "Downloading Invoice Register...")
                ps_module.download_invoices()
                
                if self.log_callback:
                    self.log_callback("INFO", "Downloading Project Earnings...")
                ps_module.download_earnings()
                
                if self.log_callback:
                    self.log_callback("INFO", "Downloading Project Expenses...")
                ps_module.download_expenses()
                
                if self.log_callback:
                    self.log_callback("INFO", "Downloading Labor Hours...")
                ps_module.download_labor_hours()
                
                # Delete default sheet
                if self.log_callback:
                    self.log_callback("INFO", "Finalizing output file...")
                ps_module.delete_default_sheet()
                
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
