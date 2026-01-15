"""
Main window for QCA Accounting Automation Tool.
"""
from typing import Dict, Optional, Any
from pathlib import Path
import inspect
import json
import time

#region agent log
DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Optional[Dict[str, object]] = None,
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

#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H3", "ui.main_window", "Before QtWidgets import", {"has_app": bool(QApplication.instance())})
except Exception as e:
    _agent_log("H3", "ui.main_window", "QtWidgets import error", {"error": str(e)})
#endregion

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QStackedWidget, QLabel, QStatusBar
)
#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H3", "ui.main_window", "After QtWidgets import", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

from PySide6.QtCore import Signal, Slot
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
from pathlib import Path

#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H2", "ui.main_window", "Before qfluentwidgets import", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

from qfluentwidgets import (
    FluentWindow,
    FluentIcon,
    setTheme,
    Theme,
    NavigationItemPosition,
    BodyLabel,
    CardWidget,
    PrimaryPushButton,
)

#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H2", "ui.main_window", "After qfluentwidgets import", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

#region agent log
try:
    _agent_log(
        "H2",
        "ui.main_window",
        "FluentWindow binding info",
        {
            "fluent_module": FluentWindow.__module__,
            "fluent_mro": [cls.__name__ for cls in FluentWindow.__mro__[:3]],
        },
    )
except Exception as e:
    _agent_log("H2", "ui.main_window", "FluentWindow binding info error", {"error": str(e)})
#endregion

#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H3", "ui.main_window", "Before ui imports", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

from ui.pages.home_page import HomePage
from ui.pages.ap_task_page import APTaskPage
from ui.pages.ar_task_page import ARTaskPage
from ui.pages.project_status_task_page import ProjectStatusTaskPage
from ui.pages.base_task_page import BaseTaskPage
from ui.pages.scheduled_tasks_page import ScheduledTasksPage
from ui.services.app_context import get_app_context
from ui.widgets.login_dialog import LoginDialog
from workers.base_worker import BaseWorker
from workers.check_login_worker import CheckLoginWorker
from workers.login_worker import LoginWorker

#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H3", "ui.main_window", "After all imports", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

class MainWindow(FluentWindow):
    """
    Main window with horizontal layout:
    - Left: Task navigation
    - Middle: Task content area
    """
    
    def __init__(self):
        """Initialize main window."""
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "enter __init__", {})
        #endregion
        try:
            app_exists = bool(QApplication.instance())
        except Exception:
            app_exists = False
        #region agent log
        _agent_log(
            "H6",
            "MainWindow.__init__",
            "before super init",
            {
                "app_exists": app_exists,
                "fluent_module": FluentWindow.__module__,
                "fluent_mro": [cls.__name__ for cls in FluentWindow.__mro__[:3]],
                "fluent_bases": [b.__module__ for b in FluentWindow.__bases__],
                "qapp_module": QApplication.__module__,
                "qapp_instance_module": type(QApplication.instance()).__module__ if QApplication.instance() else None,
                "fluent_mro_modules": [cls.__module__ for cls in FluentWindow.__mro__[:6]],
            },
        )
        #endregion
        try:
            super().__init__()
        except Exception as exc:
            #region agent log
            _agent_log(
                "H6",
                "MainWindow.__init__",
                "super init exception",
                {"error": str(exc), "app_exists": app_exists},
            )
            #endregion
            raise
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "after super init", {})
        #endregion
        
        # Set Fluent theme (required for Fluent Design appearance)
        try:
            setTheme(Theme.AUTO)  # AUTO will follow system theme
            #region agent log
            _agent_log("H6", "MainWindow.__init__", "Theme set to AUTO", {})
            #endregion
        except Exception as e:
            #region agent log
            _agent_log("H6", "MainWindow.__init__", "Failed to set theme", {"error": str(e)})
            #endregion
            # Continue without theme if it fails
        
        self.setWindowTitle("QCA Accounting Automation Tool")
        self.resize(1200, 800)
        self.setMinimumSize(900, 650)

        # Base size for proportional font scaling
        self._font_base_size = (1200, 800)
        self._font_base_pt = 11
        self._font_min_pt = 9
        self._font_max_pt = 14
        
        # Application context
        self.app_context = get_app_context()
        
        # Current executing worker
        self.current_worker: Optional[BaseWorker] = None
        
        # Task pages dictionary
        self.task_pages: Dict[str, BaseTaskPage] = {}
        
        self._setup_ui()
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "after _setup_ui", {})
        #endregion
        self._init_pages()
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "after _init_pages", {})
        #endregion
        self._connect_signals()
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "after _connect_signals", {})
        #endregion
        self._check_login_status()
        #region agent log
        _agent_log("H4", "MainWindow.__init__", "MainWindow initialized", {"pages": list(self.task_pages.keys())})
        #endregion

        # Apply initial font scaling
        self._apply_font_scale()

    def showEvent(self, event) -> None:
        """Handle window show event - set navigation to expanded state."""
        super().showEvent(event)
        # Ensure navigation is expanded after window is shown
        if self.nav_interface:
            try:
                # Try to expand navigation if it has the method
                if hasattr(self.nav_interface, 'setExpanded'):
                    self.nav_interface.setExpanded(True)
                elif hasattr(self.nav_interface, 'expanded'):
                    self.nav_interface.expanded = True
                
                # Set expand width for auto-resize
                if hasattr(self.nav_interface, 'setExpandWidth'):
                    self.nav_interface.setExpandWidth(250)
                elif hasattr(self.nav_interface, 'expandWidth'):
                    self.nav_interface.expandWidth = 250
            except Exception:
                # If setting fails, continue anyway
                pass
    
    def resizeEvent(self, event) -> None:
        """Scale font proportionally with window size for readability."""
        try:
            self._apply_font_scale()
        except Exception:
            pass
        return super().resizeEvent(event)

    def _apply_font_scale(self) -> None:
        """Apply a global font size based on current window size."""
        w = max(1, self.width())
        h = max(1, self.height())
        bw, bh = self._font_base_size
        # Use the smaller scale so text doesn't overflow vertically
        scale = min(w / max(1, bw), h / max(1, bh))
        # Clamp scale to avoid extreme sizes
        scale = max(0.85, min(scale, 1.25))
        pt = int(round(self._font_base_pt * scale))
        pt = max(self._font_min_pt, min(pt, self._font_max_pt))

        app = QApplication.instance()
        if not app:
            return

        font = app.font() or QFont()
        # Prefer Windows default UI font for readability
        if not font.family():
            font.setFamily("Segoe UI")
        font.setPointSize(pt)
        app.setFont(font)
    
    def _setup_ui(self) -> None:
        """Setup UI layout."""
        # FluentWindow already provides navigationInterface & stackedWidget
        # Access them as properties (FluentWindow API)
        self.nav_interface = self.navigationInterface
        self.stacked_widget = self.stackedWidget
        if hasattr(self.stacked_widget, "setContentsMargins"):
            self.stacked_widget.setContentsMargins(0, 0, 0, 0)
        
        # Set navigation interface to be expanded by default with auto-resize width
        if self.nav_interface:
            # Set expand width (this is the width when expanded)
            # Use a reasonable default width that will auto-resize
            if hasattr(self.nav_interface, 'setExpandWidth'):
                # Set expand width to a reasonable default (will auto-resize based on content)
                self.nav_interface.setExpandWidth(250)  # Default expanded width
            elif hasattr(self.nav_interface, 'expandWidth'):
                # If it's a property, set it directly
                self.nav_interface.expandWidth = 250
            
            # Ensure navigation is expanded by default
            if hasattr(self.nav_interface, 'setExpanded'):
                self.nav_interface.setExpanded(True)
            elif hasattr(self.nav_interface, 'expanded'):
                self.nav_interface.expanded = True
            
            # Try to set minimum width for better auto-resize behavior
            if hasattr(self.nav_interface, 'setMinimumWidth'):
                self.nav_interface.setMinimumWidth(200)
        
        # Store login worker references
        self.check_login_worker: Optional[CheckLoginWorker] = None
        self.login_worker: Optional[LoginWorker] = None
        self.login_dialog: Optional[LoginDialog] = None
    
    def _connect_signals(self) -> None:
        """Connect application context signals."""
        self.app_context.login_state_changed.connect(self._on_app_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_app_user_info_changed)
    
    def _init_pages(self) -> None:
        """Initialize all pages (home and task pages)."""
        #region agent log
        try:
            import qfluentwidgets as qfw
            sig = str(inspect.signature(self.addSubInterface))
            _agent_log(
                "H7",
                "MainWindow._init_pages",
                "addSubInterface pre-check",
                {
                    "qfw_version": getattr(qfw, "__version__", None),
                    "callable_type": str(type(self.addSubInterface)),
                    "signature": sig,
                },
            )
        except Exception as e:
            _agent_log(
                "H7",
                "MainWindow._init_pages",
                "addSubInterface pre-check error",
                {"error": str(e)},
            )
        #endregion
        
        # Create home page
        home_page = HomePage()
        home_page.setObjectName("home")  # Use objectName as identifier
        # Set up navigation callback - switch to AP page using objectName
        def navigate_to_ap():
            # Find AP page widget by objectName in stacked widget
            ap_page = None
            for i in range(self.stacked_widget.count()):
                widget = self.stacked_widget.widget(i)
                if widget and widget.objectName() == "ap":
                    ap_page = widget
                    break
            
            if ap_page:
                # Switch to AP page
                self.stacked_widget.setCurrentWidget(ap_page)
                # Update navigation interface to reflect current selection
                # FluentWindow's navigationInterface should auto-update, but we can try to set it explicitly
                try:
                    # Try to set navigation item by finding it in navigation interface
                    if hasattr(self.nav_interface, 'setCurrentItem'):
                        self.nav_interface.setCurrentItem("ap")
                except Exception:
                    # If setCurrentItem doesn't work, the stacked widget change should be enough
                    pass
        home_page._navigate_to_ap = navigate_to_ap
        
        # AR navigation
        def navigate_to_ar():
            self.stacked_widget.setCurrentWidget(self.task_pages["ar"])
            try:
                if hasattr(self.nav_interface, 'setCurrentItem'):
                    self.nav_interface.setCurrentItem("ar")
            except Exception:
                pass
        home_page._navigate_to_ar = navigate_to_ar
        home_page._trigger_login = self._on_login_clicked
        
        #region agent log
        _agent_log("H7", "MainWindow._init_pages", "About to add home page", {"home_objectName": home_page.objectName(), "runId": "post-fix"})
        #endregion
        
        # Add home page to navigation (first item)
        try:
            self.addSubInterface(
                interface=home_page,
                icon=FluentIcon.HOME,
                text="Home",
                position=NavigationItemPosition.TOP
            )
            #region agent log
            _agent_log("H7", "MainWindow._init_pages", "Home page added successfully", {"runId": "post-fix"})
            #endregion
        except Exception as e:
            #region agent log
            _agent_log("H7", "MainWindow._init_pages", "Failed to add home page", {"error": str(e), "error_type": type(e).__name__, "runId": "post-fix"})
            #endregion
            raise
        
        # Create task pages
        self.task_pages["ap"] = APTaskPage()
        self.task_pages["ar"] = ARTaskPage()
        self.task_pages["project_status"] = ProjectStatusTaskPage()
        # TODO: Add other task pages
        # self.task_pages["project_status"] = ProjectStatusTaskPage()
        # ...
        
        # Add task pages to navigation and stacked widget
        for task_id, page in self.task_pages.items():
            page.setObjectName(task_id)
            #region agent log
            try:
                call_sig = str(inspect.signature(self.addSubInterface))
            except Exception as e:
                call_sig = f"signature_error:{e}"
            _agent_log(
                "H7",
                "MainWindow._init_pages",
                "before addSubInterface",
                {
                    "task_id": task_id,
                    "page_type": type(page).__name__,
                    "signature": call_sig,
                },
            )
            #endregion
            #region agent log
            _agent_log("H7", "MainWindow._init_pages", "About to add task page", {"task_id": task_id, "objectName": page.objectName(), "runId": "post-fix"})
            #endregion
            try:
                # Load custom SVG icons from ui/icons directory
                # According to qfluentwidgets documentation: https://qfluentwidgets.com/pages/icon/#add-icon
                def load_icon(icon_name: str):
                    """Load SVG icon from ui/icons directory."""
                    icons_dir = Path(__file__).parent.parent / "ui" / "icons"
                    # Try SVG first (preferred format for qfluentwidgets)
                    svg_path = icons_dir / f"{icon_name}.svg"
                    if svg_path.exists():
                        return QIcon(str(svg_path))
                    # Fallback to PNG if SVG not found
                    png_path = icons_dir / f"{icon_name}.png"
                    if png_path.exists():
                        return QIcon(str(png_path))
                    # Fallback to FluentIcon if no custom icon found
                    return FluentIcon.DOCUMENT
                
                # Map task IDs to icon file names
                icon_map = {
                    "ap": load_icon("ap"),
                    "ar": load_icon("ar"),
                    "project_status": load_icon("project_status"),
                }
                icon = icon_map.get(task_id, FluentIcon.DOCUMENT)
                
                self.addSubInterface(
                    interface=page,
                    icon=icon,
                    text=page.task_name,
                    position=NavigationItemPosition.TOP
                )
                #region agent log
                _agent_log("H7", "MainWindow._init_pages", "Task page added successfully", {"task_id": task_id, "runId": "post-fix"})
                #endregion
            except TypeError as exc:
                #region agent log
                _agent_log(
                    "H7",
                    "MainWindow._init_pages",
                    "addSubInterface TypeError",
                    {
                        "task_id": task_id,
                        "error": str(exc),
                        "signature": call_sig,
                        "runId": "post-fix",
                    },
                )
                #endregion
                raise
            
            # Connect signals
            page.task_started.connect(self._on_task_started)
            page.task_finished.connect(self._on_task_finished)
            page.task_failed.connect(self._on_task_failed)
        #region agent log
        _agent_log("H4", "MainWindow._init_pages", "Pages initialized", {"count": len(self.task_pages) + 1})
        #endregion
        
        # Add scheduled tasks page (not a BaseTaskPage, so handle separately)
        scheduled_tasks_page = ScheduledTasksPage()
        scheduled_tasks_page.setObjectName("scheduled_tasks")
        try:
            # Load icon for scheduled tasks
            icons_dir = Path(__file__).parent.parent / "ui" / "icons"
            schedule_icon_path = icons_dir / "schedule.svg"
            if schedule_icon_path.exists():
                schedule_icon = QIcon(str(schedule_icon_path))
            else:
                schedule_icon = FluentIcon.CALENDAR
            
            self.addSubInterface(
                interface=scheduled_tasks_page,
                icon=schedule_icon,
                text="Scheduled Tasks",
                position=NavigationItemPosition.TOP
            )
        except Exception as e:
            #region agent log
            _agent_log("H7", "MainWindow._init_pages", "Failed to add scheduled tasks page", {"error": str(e)})
            #endregion
            # Continue even if icon loading fails
            try:
                self.addSubInterface(
                    interface=scheduled_tasks_page,
                    icon=FluentIcon.CALENDAR,
                    text="Scheduled Tasks",
                    position=NavigationItemPosition.TOP
                )
            except Exception:
                pass
    
    @Slot(str)
    def _on_task_started(self, task_id: str) -> None:
        """
        Handle task started signal.
        
        Args:
            task_id: Task identifier
        """
        # Disable all task pages' execute buttons
        for page in self.task_pages.values():
            page._set_execution_enabled(False)
        
        # Store current worker
        if task_id in self.task_pages:
            self.current_worker = self.task_pages[task_id].current_worker
        
        # Update status (FluentWindow doesn't have statusBar, use window title or custom widget)
        # For now, we'll update window title temporarily
        original_title = self.windowTitle()
        self.setWindowTitle(f"{original_title} - Executing: {task_id}")
        #region agent log
        _agent_log("H4", "MainWindow._on_task_started", "Task started", {"task_id": task_id})
        #endregion
    
    @Slot(str, dict)
    def _on_task_finished(self, task_id: str, result: dict) -> None:
        """
        Handle worker finished signal.
        
        Args:
            task_id: Task identifier
            result: Result dictionary
        """
        # Enable all task pages' execute buttons
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        
        # Clear current worker
        self.current_worker = None
        
        # Update window title
        self.setWindowTitle("QCA Accounting Automation Tool")
        #region agent log
        _agent_log("H4", "MainWindow._on_task_finished", "Task finished", {"task_id": task_id, "success": bool(result.get("success"))})
        #endregion
    
    @Slot(str, str)
    def _on_task_failed(self, task_id: str, error: str) -> None:
        """
        Handle worker failed signal.
        
        Args:
            task_id: Task identifier
            error: Error message
        """
        # Enable all task pages' execute buttons
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        
        # Clear current worker
        self.current_worker = None
        
        # Update window title
        self.setWindowTitle("QCA Accounting Automation Tool")
        #region agent log
        _agent_log("H4", "MainWindow._on_task_failed", "Task failed", {"task_id": task_id, "error": error})
        #endregion
    
    def _check_login_status(self) -> None:
        """Check login status using background worker."""
        # Don't check if already checking
        if self.check_login_worker and self.check_login_worker.isRunning():
            return
        
        # Create check login worker
        self.check_login_worker = CheckLoginWorker()
        
        # Connect signals
        self.check_login_worker.finished.connect(self._on_check_login_finished)
        
        # Start worker (no UI updates since navigation footer is removed)
        self.check_login_worker.start()
    
    @Slot(dict)
    def _on_check_login_finished(self, result: Dict[str, Any]) -> None:
        """
        Handle check login finished signal.
        
        Args:
            result: Result dictionary with login status
        """
        is_logged_in = result.get("is_logged_in", False)
        
        # Update app context
        user_info = {}
        if is_logged_in:
            # Try to get user email from config or result
            user_info = result.get("user_info", {})
        
        self.app_context.set_login_state(is_logged_in, user_info if user_info else None)
        
        # Update UI (will be handled by signal handler)
        self._update_login_ui()
    
    @Slot()
    def _on_login_clicked(self) -> None:
        """Handle login button click - show login dialog."""
        # Don't start if already logging in
        if self.login_worker and self.login_worker.isRunning():
            return
        
        # Show login dialog
        self.login_dialog = LoginDialog(self)
        self.login_dialog.login_requested.connect(self._start_login)
        self.login_dialog.exec()
    
    def _start_login(self) -> None:
        """Start login process."""
        # Don't start if already logging in
        if self.login_worker and self.login_worker.isRunning():
            return
        
        # Create login worker with log callback
        self.login_worker = LoginWorker(
            headless=False,
            log_callback=self._on_login_log_received
        )
        
        # Connect signals
        self.login_worker.log_signal.connect(self._on_login_log_received)
        self.login_worker.finished.connect(self._on_login_finished)
        
        # Update dialog if open
        if self.login_dialog:
            self.login_dialog._set_loading_state(True, "Opening browser for authentication...")
        
        # Start worker
        self.login_worker.start()
    
    @Slot(str, str)
    def _on_login_log_received(self, level: str, message: str) -> None:
        """
        Handle log signal from login worker.
        
        Args:
            level: Log level
            message: Log message
        """
        # Update dialog with login progress messages
        if not self.login_dialog:
            return
            
        if level == "INFO":
            # Show progress in dialog
            if "Starting" in message or "Opening" in message:
                status_msg = "Opening browser..."
                self.login_dialog._set_loading_state(True, status_msg)
            elif "completed" in message.lower() or "Verifying" in message:
                status_msg = "Verifying login..."
                self.login_dialog._set_loading_state(True, status_msg)
            else:
                # For other INFO messages, show them as status
                self.login_dialog._set_loading_state(True, message)
        elif level == "SUCCESS":
            self.login_dialog.show_success()
        elif level == "ERROR":
            self.login_dialog.show_error(message)
    
    @Slot(dict)
    def _on_login_finished(self, result: Dict[str, Any]) -> None:
        """
        Handle login finished signal.
        
        Args:
            result: Result dictionary with login status
        """
        success = result.get("success", False)
        is_logged_in = result.get("is_logged_in", False)
        
        # Update app context
        user_info = {}
        if success and is_logged_in:
            user_info = result.get("user_info", {})
        
        self.app_context.set_login_state(is_logged_in, user_info if user_info else None)
        
        # Update UI (will be handled by signal handler)
        self._update_login_ui()
        
        # Close dialog if login was successful
        if success and is_logged_in and self.login_dialog:
            # Dialog will close itself via show_success()
            pass
        
        # Re-check login status to get user email info
        if success and is_logged_in:
            self._check_login_status()
    
    def _update_login_ui(self) -> None:
        """Update login UI based on app context state."""
        # Navigation footer removed, no UI to update
        # Login state is managed via AppContext and displayed in HomePage
        pass
    
    @Slot(bool)
    def _on_app_login_state_changed(self, is_logged_in: bool) -> None:
        """
        Handle app context login state changed signal.
        
        Args:
            is_logged_in: Whether user is logged in
        """
        self._update_login_ui()
    
    @Slot(dict)
    def _on_app_user_info_changed(self, user_info: dict) -> None:
        """
        Handle app context user info changed signal.
        
        Args:
            user_info: User information dictionary
        """
        self._update_login_ui()
