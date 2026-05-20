"""
Unified entry point for QCA Accounting Automation Tool.

Supports both CLI mode (with command-line arguments) and UI mode (without arguments).
- CLI mode: Execute tasks based on command-line arguments
- UI mode: Launch the graphical user interface

All update checking logic has been removed for simplified startup.
"""
from getpass import getpass
import sys
import json
import time
import argparse
from pathlib import Path
from typing import Optional
from PySide6.QtCore import QTimer
from tools.runtime_logger import RuntimeLogger
import tools.util as util
import tools.login as login
import os
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
    
    progress_dialog = None
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        from PySide6.QtCore import Qt, QTimer
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
        _agent_log("H5", "main.launch_ui", "QApplication created", {"app": bool(app), "instance": bool(QApplication.instance())})
        #endregion
        #region agent log
        _agent_log("H5", "main.launch_ui", "After set app properties", {})
        #endregion

        # Show a lightweight startup progress UI to improve perceived startup performance.
        try:
            from ui.widgets.startup_progress_dialog import StartupProgressDialog

            progress_dialog = StartupProgressDialog()
            progress_dialog.show()
            progress_dialog.set_progress(5, "Initializing application...")
            app.processEvents()
        except Exception:
            # Startup progress is best-effort; never block startup if it fails.
            progress_dialog = None

        def _report_startup(percent: int, text: str) -> None:
            """Update startup progress UI (best-effort)."""
            if progress_dialog:
                try:
                    progress_dialog.set_progress(percent, text)
                    app.processEvents()
                except Exception:
                    pass
        
        # Import MainWindow after QApplication is created
        #region agent log
        _agent_log("H3", "main.launch_ui", "Before importing MainWindow", {"has_app": bool(QApplication.instance())})
        #endregion
        _report_startup(15, "Loading UI modules...")
        app.processEvents()
        
        from ui.main_window import MainWindow
        app.processEvents()
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
        _report_startup(35, "Building main window...")
        app.processEvents()
        
        # Create main window (this may take time)
        window = MainWindow(startup_progress=_report_startup)
        app.processEvents()
        
        #region agent log
        _agent_log("H4", "main.launch_ui", "MainWindow created", {"has_app": bool(QApplication.instance())})
        #endregion
        _report_startup(90, "Finalizing...")
        app.processEvents()
        
        # Show main window
        window.show()
        app.processEvents()

        # Only close PyInstaller splash in frozen executable
        if getattr(sys, "frozen", False):
            try:
                import pyi_splash
                pyi_splash.close()
            except Exception:
                pass

        #region agent log
        _agent_log("H5", "main.launch_ui", "Window shown", {})
        #endregion

        # Close the startup progress UI shortly after the main window is visible.
        if progress_dialog:
            _report_startup(100, "Ready")
            QTimer.singleShot(250, progress_dialog.close)
        
        # Run event loop
        #region agent log
        _agent_log("H5", "main.launch_ui", "About to exec event loop", {})
        #endregion
        return app.exec()
    except Exception as exc:
        #region agent log
        _agent_log("H5", "main.launch_ui", "Main exception", {"error": str(exc), "type": type(exc).__name__, "has_app": bool(QApplication.instance())})
        #endregion
        # Ensure startup progress UI is closed if we fail during startup.
        try:
            if progress_dialog:
                progress_dialog.close()
        except Exception:
            pass
        
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

    '''    
    import core.revenue_accrual as ra
    login.sso_login()
    ra.main()

    import core.project_status as ps
    #login.sso_login()
    ps.main()
    
    import core.daily_received as dr
    dr.main()
    '''

    util.init_workdir()

    exit_code = main()
    if exit_code == 0 and not any(
        arg in sys.argv for arg in ["--ap", "--ar", "--project_status", "--bridge", 
                                    "--shipping", "--daily_received", "--on_call", "--init_schedule"]
    ):
        # Only print completion message for CLI mode
        pass
    sys.exit(exit_code)
    '''

    #raw_projects = "Q-08571,Q-08571_CO1,Q-08690J-D2,Q-08771,Q-08835,Q-08836C,Q-08836D,Q-08836E,Q-08836F,Q-08836G,Q-08836I,Q-08836K,Q-08884,Q-08888,Q-08889,Q-08891,Q-08897,Q-08899,Q-08900,Q-08901,Q-08904,Q-08905,Q-08907,Q-08908,Q-08909,Q-08911,Q-08913,Q-08914,Q-08915,Q-08916,Q-08917,Q-08918,Q-08920,Q-08923,Q-08929,Q-08930,Q-08931,Q-08933,Q-08934,Q-08936,Q-08938,Q-08939,Q-08940,Q-08942,Q-08943,Q-08944,Q-08945,Q-08946,Q-08948,Q-08949,Q-08950,Q-08954,Q-08955,Q-08962,Q-08963,Q-08964,Q-08965,Q-08967,Q-08968,Q-08971,Q-08974,Q-08975,Q-08976,Q-08977,Q-08978,Q-08979,Q-08982,Q-08984,Q-08985,Q-08986,Q-08988,Q-08989,Q-08991,Q-08993,Q-08994,Q-08996,Q-08998,Q-08999,Q-09001,Q-09002,Q-09007,Q-09009,Q-09010,Q-09012,Q-09013,Q-09015,Q-09016,Q-09022,Q-09023,Q-09024,Q-09025,Q-09026,Q-09028,Q-09029,Q-09033,Q-09036,Q-09037,Q-09040,Q-09045,Q-09046,Q-09047,Q-09048,Q-09049,Q-09050,Q-09051,Q-09052,Q-09053,Q-09057,Q-09058,Q-09059,Q-09060,Q-09062,Q-09064,Q-09065,Q-09067,Q-09068,Q-09070,Q-09071,Q-09072,Q-09072_01,Q-09075,Q-09076,Q-09077,Q-09078,Q-09079,Q-09080,Q-09082,Q-09083,Q-09084,Q-09085,Q-09086,Q-09087,Q-09090,Q-09091,Q-09092,Q-09096,Q-09100,Q-09102,Q-09103,Q-09104,Q-09105,Q-09106,Q-09108,Q-09109,Q-09111,Q-09112,Q-09114,Q-09116,Q-09117,Q-09120,Q-09122,Q-09123,Q-09126,Q-09127,Q-09128,Q-09130,Q-09131,Q-09132,Q-09134,Q-09135,Q-09136,Q-09137,Q-09140,Q-09145,Q-09146,Q-09147,Q-09148,Q-09149,Q-09151,Q-09152,Q-09153,Q-09155,Q-09157,Q-09159,Q-09160,Q-09161,Q-09163,Q-09165,Q-09166,Q-09168,Q-09169,Q-09173,Q-09174,Q-09175,Q-09176,Q-09178,Q-09181,Q-09185,Q-09186,Q-09188,Q-09190,Q-09191,Q-09196,Q-09198,Q-09199,Q-09203,Q-09213,Q-09217,Q-09218,Q-09220,Q-09222,Q-09223,Q-09227,Q-09228,Q-09230,Q-09231,Q-09233,Q-09235,Q-09239,Q-09240,Q-09242,Q-09244,Q-09246,Q-09247,Q-09249,Q-09250,Q-09252,Q-09253,Q-09254,Q-09256,Q-09257,Q-09258,Q-09261,Q-09262,Q-09264,Q-09265,Q-09266,Q-09268,Q-09271,Q-09273,Q-09274,Q-09276,Q-09282,Q-09283,Q-09285,Q-09287,Q-09289,Q-09292,Q-09293,Q-09296,Q-09298,Q-09299,Q-09300,Q-09301,Q-09302,Q-09305,Q-09306,Q-09307,Q-09310,Q-09311,Q-09313,Q-09318,Q-09328,Q-09329,Q-09333,Q-09334,Q-09337,Q-09346,Q-2025-CB,Q-2025-NBT,Q-2026-CB,Q-2026-NBT,Q-5715NBT-WETCOM,Q-6585,Q-7075,Q-7118,Q-7173,Q-7210,Q-7210-CO1,Q-7224,Q-7265NBT-C8,Q-7265NBT-PTOS,Q-7295,Q-7301,Q-7316-2025,Q-7387,Q-7476,Q-7589,Q-7681,Q-7687,Q-7728,Q-7735,Q-7740,Q-7742,Q-7754,Q-7768,Q-7779,Q-7780,Q-7791,Q-7796,Q-7813-SR110,Q-7833,Q-7849,Q-7851,Q-7869,Q-7888,Q-7897,Q-7921,Q-7965,Q-7987,Q-7999,Q-8021,Q-8027,Q-8039,Q-8039_CO1,Q-8052,Q-8056,Q-8075,Q-8083,Q-8083_COMMISSIONING,Q-8108,Q-8132,Q-8139,Q-8146,Q-8147,Q-8152,Q-8156,Q-8160,Q-8172,Q-8172-CO1,Q-8172-TT2,Q-8197,Q-8208,Q-8209,Q-8223DP,Q-8223SR1,Q-8223SR2,Q-8246,Q-8249,Q-8252,Q-8262,Q-8273,Q-8279,Q-8284,Q-8301,Q-8315SR5,Q-8315SR6,Q-8331-ALP,Q-8331-HH,Q-8331-HH_T&M,Q-8335,Q-8341,Q-8355,Q-8357,Q-8358,Q-8365,Q-8367,Q-8372,Q-8386-2026,Q-8386-2026-01,Q-8387,Q-8390,Q-8417,Q-8430,Q-8446,Q-8464,Q-8476,Q-8480,Q-8481,Q-8489,Q-8491,Q-8494,Q-8499,Q-8510,Q-8511,Q-8520,Q-8529,Q-8537,Q-8540,Q-8541,Q-8544,Q-8544_COMMISSIONING,Q-8548,Q-8556,Q-8561,Q-8566,Q-8575,Q-8584,Q-8592,Q-8595,Q-8599,Q-8600,Q-8604,Q-8605,Q-8614,Q-8616,Q-8625,Q-8628,Q-8630,Q-8632,Q-8636,Q-8640,Q-8646,Q-8651,Q-8655,Q-8656,Q-8662,Q-8665,Q-8669,Q-8672,Q-8674,Q-8682,Q-8683,Q-8689,Q-8690_TRAVEL,Q-8690A,Q-8690A-CONTINGENCY,Q-8690B,Q-8690B-CONTINGENCY,Q-8690C,Q-8690F,Q-8690G,Q-8690G-CONTINGENCY,Q-8692,Q-8697,Q-8700,Q-8701,Q-8709,Q-8710,Q-8711,Q-8714,Q-8715,Q-8720,Q-8721,Q-8725,Q-8727,Q-8727_CO1,Q-8735,Q-8736,Q-8737,Q-8739,Q-8741,Q-8742,Q-8743,Q-8745,Q-8750,Q-8752,Q-8753,Q-8754,Q-8756,Q-8758,Q-8759,Q-8762,Q-8763,Q-8765,Q-8770,Q-8772,Q-8772_COMMISSIONING,Q-8775,Q-8776,Q-8779,Q-8780,Q-8781,Q-8782,Q-8784,Q-8785,Q-8789,Q-8797,Q-8799,Q-8802,Q-8805,Q-8806,Q-8808,Q-8809,Q-8813,Q-8815,Q-8821,Q-8827,Q-8829,Q-8832,Q-8834,Q-8836,Q-8836A,Q-8836B,Q-8837,Q-8838,Q-8843,Q-8848,Q-8850,Q-8852,Q-8854,Q-8855,Q-8856,Q-8858,Q-8861,Q-8864,Q-8865,Q-8868,Q-8872,Q-8874,Q-8875,Q-8876,Q-8878,Q-8879A1,Q-8879A2,Q-8879A3,Q-8879D,Q-8881,Q-8882,Q-8886,Q-9999,Q-BCMEA-2026A,Q-BCMEA-AUG-2025,Q-BCRTC-AUG25-01,Q-BCRTC-MAR26-01,Q-BCRTC-MAR26-02,QC-8041,QC-8505,QC-8608,Q-CRD2025,Q-CRD2026,Q-CSX-2026D,Q-DD2025,Q-DPWPARTS-AUG25-01,Q-DUBOIS2025,Q-FGT2025,Q-FGT2026,Q-GCT2025,Q-GCT2026,Q-GCTPARTS-APR25-01,Q-GCTPARTS-FEB26-01,Q-GCTPARTS-JAN26-01,Q-GCTPARTS-MAR25-01,Q-GCTPARTS-MAR26-01,Q-GCTPARTS-MAY25-01,Q-GFL2025,Q-GFL2026,Q-GFLPARTS-AUG2025,Q-NBT2025,Q-NBT2025-24/7,Q-NBT2025-UN,Q-NBT2026,Q-NBT2026-24/7,Q-NBT2026-UN,Q-NBT-CAR,Q-NBTPARTS-AUG-25-01,Q-NBTPARTS-FEB26-01,Q-NBTPARTS-FEB26-02,Q-NBTPARTS-FEB26-03,Q-NBTPARTS-FEB26-04,Q-NBTPARTS-JAN24-04,Q-NBTPARTS-JAN26-01,Q-NBTPARTS-JAN26-02,Q-NBTPARTS-JULY25-01,Q-NBTPARTS-JULY25-02,Q-NBTPARTS-JULY25-03,Q-NBTPARTS-JULY25-04,Q-NBTPARTS-JUNE25-01,Q-NBTPARTS-JUNE25-02,Q-NBTPARTS-MAR25-01,Q-NBTPARTS-MAR25-02,Q-NBTPARTS-MAY25-01,Q-NBTPARTS-NOV25-01,Q-NBTPARTS-NOV25-02,Q-NBTPARTS-NOV25-03,Q-NBTPARTS-OCT25-01,Q-NBTPARTS-SEP-25-01,Q-NBTPARTS-SEP25-02,Q-NBTPARTS-SEPT25-03,Q-PEM2025,Q-PEM2026,Q-RMOW2025,Q-RMOW2026,Q-SYN2025,Q-SYN2026,Q-TFN2025,Q-TFN2026,Q-TL2025,Q-USED-2025-ATK,Q-USED-2025-BJP,Q-USED-2025-EPD,Q-USED-2025-FN2,Q-USED-2025-LHG,Q-USED-2025-MCL,Q-USED-2025-NLO,Q-USED-2025-SF2,Q-USED-2025-SF2-2,Q-USED-2025-UPG,Q-USED-2025-WJC,Q-USED-2025-WJC2,Q-USED-2026-DJC,Q-USED-2026-MK,Q-WC2025,Q-08892,Q-08898,Q-08930_IGNITION,Q-09205,Q-7621,Q-8190,Q-8377,Q-8572,Q-8615,Q-8617,Q-8648,Q-8680,Q-8708,Q-8846,Q-8851,Q-8869,Q-8873,Q-8910,Q-AVC2025,Q-GCTPARTS-FEB25-01,Q-GFL2024,Q-NBTPARTS-FEB25-01,Q-8243,Q-09330,Q-09332,Q-09335,Q-09385,Q-8483,Q-CERT-2026,Q-GCTPARTS-MAR26-02,Q-NBTPARTS-APR26-02,Q-USED-2026-BYL,Q-USED-2026-GSP,Q-08690J-D3,Q-09187,Q-09208,Q-09339,Q-09340,Q-09343,Q-09351,Q-09356,Q-09357,Q-09358,Q-09359,Q-09361,Q-09364,Q-09365,Q-09370,Q-09373,Q-09374,Q-09384,Q-09391,Q-09393,Q-DD2026,Q-DUBOIS2026,Q-NBT2026-UN-04APR,Q-7316-2026,Q-09349,Q-GCTPARTS-APR26-01,Q-NBTPARTS-APR26-01,Q-NBTPARTS-APR26-03,Q-NBTPARTS-MAY26-01"
    raw_projects = "Q-7316-2026,Q-09349,Q-GCTPARTS-APR26-01,Q-NBTPARTS-APR26-01,Q-NBTPARTS-APR26-03,Q-NBTPARTS-MAY26-01"
    projects = raw_projects.split(",")
    #login.sso_login()
    header = util.set_headers()
    util.save_new_search_options(header=header, projects=projects, saveName="jay_test", isPublic = False)
    '''