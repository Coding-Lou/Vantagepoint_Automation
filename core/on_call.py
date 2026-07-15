import os
import win32com.client
import pythoncom
import tools.form_notifier as form_notifier
import tools.util as util
from datetime import datetime
from zoneinfo import ZoneInfo
import time

def run_oncall_task(vendorName: str):
    today = datetime.now(ZoneInfo("America/Vancouver"))
    YEAR = today.strftime("%Y")
    DATE = today.strftime("%Y-%m-%d")
    EXECUTOR = util.get_config(["Schedule_TASK_NOTIFICATION", "EXECUTOR"])
    FOLDER_PATH = util.get_config(["ON_CALL", "FOLDER_PATH"])
    MACRO_NAME = util.get_config(["ON_CALL", "MACRO_NAME"])
    excel_name = f"{YEAR} On-Call calendar_{vendorName}.xlsm"

    # Construct full path to Excel file
    excel_path = os.path.join(FOLDER_PATH, "On-Call calendar", excel_name)
    if not os.path.exists(excel_path):
        msg = f"Excel file not found: {excel_name}"
        form_notifier.send_task_execution_status(
            task_name= f"On Call Task {DATE}",
            executor= EXECUTOR,
            status= "Failed",
            message= f"On Call task failed with error: {msg}"
        )
        return

    excel = None
    wb = None
    status = "SUCCESS"
    error_msg = ""

    try:
        pythoncom.CoInitialize()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(excel_path)

        # Run macro (fully qualified)
        excel.Application.Run(f"'{wb.Name}'!{MACRO_NAME}")

        # Read execution status
        try:
            status = wb.Names("PY_STATUS").RefersTo
            status = status.replace('=', '').replace('"', '')

            if status == "FAILED":
                error_msg = wb.Names("PY_ERROR").RefersTo
                error_msg = error_msg.replace('=', '').replace('"', '')
            else:
                error_msg = None

        except Exception:
            status = "SUCCESS"
            error_msg = None

    except Exception as e:
        status = "FAILED"
        error_msg = str(e)
    
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
    
        if excel is not None:
            excel.Quit()
    
        pythoncom.CoUninitialize()

    # Prepare notification message
    message = f"On-call VBA execution for '{excel_name}' finished.\nStatus: {status}"
    if error_msg:
        message += f"\nError: {error_msg}"

    # Send notification using form_notifier
    form_notifier.send_task_execution_status(
            task_name= f"On Call Task {DATE}",
            executor= EXECUTOR,
            status= status,
            message= message
        )
    
    return

