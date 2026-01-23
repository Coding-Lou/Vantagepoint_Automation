"""
Unified entry point for QCA Accounting Automation Tool.

Supports both CLI mode (with command-line arguments) and UI mode (without arguments).
- CLI mode: Execute tasks based on command-line arguments
- UI mode: Launch the graphical user interface

All update checking logic has been removed for simplified startup.
"""
import sys
import json
import time
import argparse
from pathlib import Path
from typing import Optional

from tools.runtime_logger import RuntimeLogger
import tools.util as util
import tools.login as login
# Core modules and schedule_task are imported lazily in run_task_by_flag to avoid import errors

# =========================
# Set up logger
# =========================
BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
log_file = LOGS_DIR / "runtime.log"

DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Optional[dict] = None,
    run_id: str = "pre-fix",
) -> None:
    """Append a compact NDJSON log for debug workflow."""
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


def run_task_by_flag(args) -> bool:
    """
    Execute task based on command-line arguments.
    
    Core modules are imported lazily here to avoid import errors at startup.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        True if a task was executed, False otherwise
    """
    LOGIN = util.check_login()

    if args.ap:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.ap as ap
        ap.main()
    elif args.ar:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.ar as ar
        ar.main()
    elif args.project_status:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.project_status as project_status
        project_status.main()
    elif args.bridge:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.bridge_report as bridge_report
        bridge_report.main()
    elif args.shipping:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.shipping_monitor as shipping_monitor
        shipping_monitor.main()
    elif args.daily_received:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        import core.daily_received as daily_received
        daily_received.main()
    elif args.on_call:
        import core.on_call as on_call
        on_call.run_oncall_task(vendorName=args.name)
    elif args.init_schedule:
        import tools.schedule_task as schedule_task
        schedule_task.main()
    else:
        return False

    return True


def launch_ui() -> int:
    """
    Launch the graphical user interface.
    
    Returns:
        Exit code from QApplication.exec()
    """
    #region agent log
    _agent_log("H1", "main.launch_ui", "UI launch starting", {"argv": sys.argv})
    #endregion
    
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QGuiApplication, QFont
        
        #region agent log
        _agent_log(
            "H1",
            "main.launch_ui",
            "Before setting High DPI attributes",
            {"has_app": bool(QApplication.instance())},
        )
        #endregion
        
        # Set High DPI scaling policy (Qt 6)
        try:
            QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
                Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
            )
            #region agent log
            _agent_log("H1", "main.launch_ui", "High DPI rounding policy set", {})
            #endregion
        except Exception as e:
            #region agent log
            _agent_log("H1", "main.launch_ui", "High DPI rounding policy not available", {"error": str(e)})
            #endregion
            pass
        
        #region agent log
        _agent_log("H5", "main.launch_ui", "About to create QApplication", {"has_app": bool(QApplication.instance())})
        #endregion
        
        # Create application (must be before any QWidget creation)
        app = QApplication(sys.argv)
        #region agent log
        _agent_log("H5", "main.launch_ui", "QApplication created", {"app": bool(app), "instance": bool(QApplication.instance())})
        #endregion
        
        # Set application properties
        app.setApplicationName("QCA Accounting Automation Tool")
        app.setOrganizationName("QCA Systems Ltd")

        # Set default font
        try:
            base_font = QFont("Segoe UI", 11)
            app.setFont(base_font)
        except Exception:
            pass
        #region agent log
        _agent_log("H5", "main.launch_ui", "After set app properties", {})
        #endregion
        
        # Import MainWindow after QApplication is created
        #region agent log
        _agent_log("H3", "main.launch_ui", "Before importing MainWindow", {"has_app": bool(QApplication.instance())})
        #endregion
        from ui.main_window import MainWindow
        #region agent log
        _agent_log(
            "H2",
            "main.launch_ui",
            "MainWindow imported",
            {
                "mro_modules": [cls.__module__ for cls in MainWindow.__mro__[:6]],
            },
        )
        #endregion
        
        # Check Qt binding consistency
        mro_modules = [cls.__module__ for cls in MainWindow.__mro__[:6]]
        app_binding = QApplication.__module__.split(".")[0]
        mw_binding = "PyQt5" if any("PyQt5" in m for m in mro_modules) else ("PySide6" if any("PySide6" in m for m in mro_modules) else "unknown")
        if mw_binding != "unknown" and mw_binding != app_binding:
            msg = (
                f"Qt binding mismatch: app uses {app_binding}, MainWindow uses {mw_binding}. "
                "Install the PySide6 build of Fluent Widgets (pip uninstall qfluentwidgets; "
                "pip install PySide6-Fluent-Widgets) or switch the app to PyQt5."
            )
            #region agent log
            _agent_log(
                "H2",
                "main.launch_ui",
                "Qt binding mismatch detected",
                {"app_binding": app_binding, "mw_binding": mw_binding},
            )
            #endregion
            raise RuntimeError(msg)
        #region agent log
        _agent_log("H3", "main.launch_ui", "After importing MainWindow", {"has_app": bool(QApplication.instance())})
        #endregion
        
        # Create and show main window
        #region agent log
        _agent_log("H4", "main.launch_ui", "Before creating MainWindow instance", {"has_app": bool(QApplication.instance())})
        #endregion
        window = MainWindow()
        #region agent log
        _agent_log("H4", "main.launch_ui", "MainWindow created", {"has_app": bool(QApplication.instance())})
        #endregion
        window.show()
        #region agent log
        _agent_log("H5", "main.launch_ui", "Window shown", {})
        #endregion
        
        # Run event loop
        #region agent log
        _agent_log("H5", "main.launch_ui", "About to exec event loop", {})
        #endregion
        return app.exec()
    except Exception as exc:
        #region agent log
        _agent_log("H5", "main.launch_ui", "Main exception", {"error": str(exc), "type": type(exc).__name__, "has_app": bool(QApplication.instance())})
        #endregion
        # Print error to console for debugging
        import traceback
        error_msg = f"Application error: {str(exc)}\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        # Try to show error dialog if QApplication exists
        try:
            from PySide6.QtWidgets import QApplication, QMessageBox
            if QApplication.instance():
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Icon.Critical)
                msg_box.setWindowTitle("Application Error")
                msg_box.setText(f"An error occurred:\n{str(exc)}")
                msg_box.setDetailedText(traceback.format_exc())
                msg_box.exec()
        except Exception:
            pass
        raise


def main() -> int:
    """
    Main entry point for the application.
    
    Behavior:
    - If command-line arguments are provided, execute the corresponding task (CLI mode)
    - If no arguments are provided, launch the UI (UI mode)
    
    Returns:
        Exit code (0 for success, non-zero for errors)
    """
    parser = argparse.ArgumentParser(description="Vantagepoint Automation Tool")
    parser.add_argument("--ap", action="store_true", help="Run AP task")
    parser.add_argument("--ar", action="store_true", help="Run AR task")
    parser.add_argument("--project_status", action="store_true", help="Run Project Status task")
    parser.add_argument("--bridge", action="store_true", help="Run Bridge Report task")
    parser.add_argument("--shipping", action="store_true", help="Run Shipping Monitor task")
    parser.add_argument("--daily_received", action="store_true", help="Run Daily Receiving task")
    parser.add_argument("--on_call", action="store_true", help="Run On Call task")
    parser.add_argument("--name", type=str, help="Vendor name for on-call task (e.g. Pembina)")
    parser.add_argument("--init_schedule", action="store_true", help="Initialize scheduled tasks")

    args = parser.parse_args()

    # Check if any task flag is set
    has_task_flag = (
        args.ap or args.ar or args.project_status or args.bridge or
        args.shipping or args.daily_received or args.on_call or args.init_schedule
    )

    # CLI mode: Execute task based on arguments
    if has_task_flag:
        with RuntimeLogger(log_file) as logger:
            if run_task_by_flag(args):
                return 0
            else:
                print("No valid task specified.")
                return 1

    # UI mode: Launch graphical interface (no arguments provided)
    return launch_ui()


if __name__ == "__main__":
    exit_code = main()
    if exit_code == 0 and not any(
        arg in sys.argv for arg in ["--ap", "--ar", "--project_status", "--bridge", 
                                    "--shipping", "--daily_received", "--on_call", "--init_schedule"]
    ):
        # Only print completion message for CLI mode
        pass
    sys.exit(exit_code)
