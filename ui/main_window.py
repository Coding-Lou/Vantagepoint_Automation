"""
Main window for QCA Accounting Automation Tool.
"""
from typing import Dict, Optional, Any
from pathlib import Path
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
        pass
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
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QFrame,
)
#region agent log
try:
    from PySide6.QtWidgets import QApplication
    _agent_log("H3", "ui.main_window", "After QtWidgets import", {"has_app": bool(QApplication.instance())})
except Exception:
    pass
#endregion

from PySide6.QtCore import Signal, Slot, Qt, QRect, QEvent, QTimer
from PySide6.QtGui import QFont, QIcon, QPainter, QBrush, QColor, QPalette, QShowEvent

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
    NavigationWidget,
    isDarkTheme,
    MessageBox,
)

from qfluentwidgets import FluentIcon as FIF
from qframelesswindow import TitleBar

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
from ui.pages.revenue_accrual_task_page import RevenueAccrualTaskPage
from ui.pages.statement_check_task_page import StatementCheckTaskPage
from ui.pages.base_task_page import BaseTaskPage
from ui.pages.scheduled_tasks_page import ScheduledTasksPage
from ui.services.app_context import get_app_context
from ui.widgets.login_dialog import LoginDialog
from ui.utils.theme_colors import ThemeColors
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


class CustomAvatarWidget(NavigationWidget):
    """Custom avatar widget for navigation bar bottom"""
    
    def __init__(self, parent=None):
        super().__init__(isSelectable=False, parent=parent)
        self.avatar_image = None
        self.initials = ""
        self.user_email = None
        # Get AppContext to listen for login state changes
        from ui.services.app_context import get_app_context
        self.app_context = get_app_context()
        self.app_context.login_state_changed.connect(self._on_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_user_info_changed)
        self._load_avatar()
    
    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        """Handle login state changed signal."""
        self._load_avatar()
        self.update()
    
    def _on_user_info_changed(self, user_info: dict) -> None:
        """Handle user info changed signal."""
        self._load_avatar()
        self.update()
    
    def _load_avatar(self):
        """Load avatar image or generate initials from AppContext."""
        # Get user email from AppContext first, fallback to direct check
        user_email = self.app_context.user_email
        if not user_email:
            # Fallback: check directly if AppContext doesn't have it yet
            import tools.util as util
            user_email = util.check_login()
            # Update AppContext if we got email from direct check
            if user_email:
                user_info = {"EMail": user_email, "email": user_email}
                self.app_context.set_login_state(True, user_info)
        
        self.user_email = user_email
        
        if user_email:
            # Extract initials from email
            email_prefix = user_email.split("@")[0]
            name_parts = email_prefix.split(".")
            if len(name_parts) >= 2:
                self.initials = (name_parts[0][0] + name_parts[1][0]).upper()
            else:
                self.initials = email_prefix[:2].upper() if len(email_prefix) >= 2 else email_prefix[0].upper()
        else:
            self.initials = "?"
    
    def paintEvent(self, e):
        """Paint the avatar widget"""
        painter = QPainter(self)
        painter.setRenderHints(
            QPainter.SmoothPixmapTransform | QPainter.Antialiasing
        )
        
        painter.setPen(Qt.NoPen)
        
        if self.isPressed:
            painter.setOpacity(0.7)
        
        # Draw background
        if self.isEnter:
            c = 255 if isDarkTheme() else 0
            painter.setBrush(QColor(c, c, c, 10))
            painter.drawRoundedRect(self.rect(), 5, 5)
        
        # Draw avatar circle
        if self.isCompacted:
            avatar_rect = QRect(12, 6, 24, 24)
        else:
            # Draw avatar circle with initials
            avatar_rect = QRect(8, 6, 24, 24)
        
        painter.setBrush(QBrush(QColor(100, 150, 200) if not isDarkTheme() else QColor(70, 120, 170)))
        painter.drawEllipse(avatar_rect)
            
        # Draw initials text
        painter.setPen(Qt.white)
        font = QFont('Segoe UI')
        font.setPixelSize(12)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(avatar_rect, Qt.AlignCenter, self.initials)
        
        # Draw name text (email prefix)
        if not self.isCompacted:
            if self.user_email:
                email_prefix = self.user_email.split("@")[0]
                painter.setPen(Qt.white if isDarkTheme() else Qt.black)
                font = QFont('Segoe UI')
                font.setPixelSize(14)
                painter.setFont(font)
                painter.drawText(QRect(44, 0, 255, 36), Qt.AlignVCenter, email_prefix)


class CustomTitleBar(TitleBar):
    """
    Custom title bar with icon and title on the left.
    Implementation adapted from PyQt-Fluent-Widgets navigation2/demo.py
    """
    
    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TranslucentBackground)
        # Create labels
        self.iconLabel = QLabel(self)
        self.titleLabel = QLabel(self)

        self.iconLabel.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.titleLabel.setAttribute(Qt.WA_TransparentForMouseEvents)
        
        # Configure label styling
        self.titleLabel.setObjectName('titleLabel')
        # Explicitly set font to ensure visibility
        self.titleLabel.setFont(QFont("Segoe UI", 12))
        self.iconLabel.setFixedSize(18, 18)
        
        # Insert widgets into the layout
        # Index 0 is often reserved or empty in TitleBar, we insert at the start
        self.hBoxLayout.insertSpacing(0, 20)
        self.hBoxLayout.insertWidget(1, self.iconLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.hBoxLayout.insertWidget(2, self.titleLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        
        # Connect signals
        self.window().windowIconChanged.connect(self.setIcon)
        self.window().windowTitleChanged.connect(self.setTitle)
    
    def setTitle(self, title):
        self.titleLabel.setText(title)
        self.titleLabel.adjustSize()
    
    def setIcon(self, icon):
        self.iconLabel.setPixmap(QIcon(icon).pixmap(18, 18))


# Import SettingsPage
from ui.widgets.settings_page import SettingsPage

class SettingsWidget(QFrame):
    """Settings widget wrapper"""
    
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setObjectName("settings")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Create settings page
        self.settings_page = SettingsPage(self)
        self.layout.addWidget(self.settings_page)


class MainWindow(FluentWindow):
    """
    Main window with horizontal layout:
    - Left: Task navigation
    - Middle: Task content area
    """
    
    def __init__(self, startup_progress=None):
        """
        Initialize main window.
        
        Args:
            startup_progress: Optional callback function(progress: int, message: str) for startup progress reporting
        """
        #region agent log
        _agent_log("H6", "MainWindow.__init__", "enter __init__", {})
        #endregion
        self._startup_progress = startup_progress
        
        try:
            super().__init__()
        except Exception as exc:
            #region agent log
            _agent_log("H6", "MainWindow.__init__", "super init exception", {"error": str(exc)})
            #endregion
            raise
        
        # 1. Initialize UI Layout (Size, TitleBar)
        self._report_progress(40, "Initializing window...")
        self._init_window()
        
        # 2. Application context & State
        self._report_progress(50, "Loading application context...")
        self.app_context = get_app_context()
        self.current_worker: Optional[BaseWorker] = None
        self.task_pages: Dict[str, BaseTaskPage] = {}
        self._initialization_complete = False  # Flag to track initialization status
        
        # Base size for proportional font scaling
        self._font_base_size = (1200, 800)
        self._font_base_pt = 11
        self._font_min_pt = 9
        self._font_max_pt = 14
        
        # 3. Setup Navigation & Pages
        self._report_progress(60, "Setting up navigation...")
        self._setup_ui()
        self._report_progress(70, "Initializing pages...")
        self._init_pages()
        self._report_progress(80, "Connecting signals...")
        self._connect_signals()
        self._report_progress(85, "Checking login status...")
        self._check_login_status()
        
        # Apply initial font scaling
        #self._apply_font_scale()
        
        # Mark initialization as complete
        self._initialization_complete = True
        
        self._report_progress(90, "Ready")
        
        #region agent log
        _agent_log("H4", "MainWindow.__init__", "MainWindow initialized", {"pages": list(self.task_pages.keys())})
        #endregion
    
    def _report_progress(self, progress: int, message: str) -> None:
        """
        Report startup progress if callback is available.
        
        Args:
            progress: Progress value (0-100)
            message: Progress message
        """
        if self._startup_progress:
            try:
                self._startup_progress(progress, message)
            except Exception:
                # Silently ignore errors in progress reporting to avoid breaking startup
                pass

    def center_window(self):
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_center_x = screen_geometry.width() // 2
        screen_center_y = screen_geometry.height() // 2

        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())

        self.move(window_geometry.topLeft())

    def _init_window(self):
        """Initialize basic window properties and custom TitleBar."""
        self.setTitleBar(CustomTitleBar(self))
        self.setWindowTitle("QCA Accounting Automation Tool")
        self.resize(1200, 800)
        self.setMinimumSize(900, 650)
        self.center_window()
        
        # Set Fluent theme
        try:
            setTheme(Theme.LIGHT)
        except Exception as e:
            _agent_log("H6", "MainWindow.__init__", "Failed to set theme", {"error": str(e)})

    def _setup_ui(self) -> None:
        """Configure the navigation interface settings."""
        # Clean configuration for Navigation 2 style
        self.navigationInterface.setExpandWidth(280)
        self.navigationInterface.setMinimumExpandWidth(300)
        
        self.stackedWidget.setContentsMargins(0, 0, 0, 0)
        
        # Apply theme-aware background to stacked widget (with error handling)
        try:
            self._apply_stacked_widget_style()
        except Exception:
            # Silently ignore errors during initialization
            pass
        
        # Keep default paint behavior to avoid visual artifacts
        
        # Store login worker references
        self.check_login_worker: Optional[CheckLoginWorker] = None
        self.login_worker: Optional[LoginWorker] = None
        self.login_dialog: Optional[LoginDialog] = None
        
        # Store bottom navigation widgets
        self.avatar_widget: Optional[CustomAvatarWidget] = None
        
        # Cache navigation width to avoid frequent queries during page switching
        self._cached_nav_width = 0
        self._is_switching_page = False
        
        # Connect to navigation display mode changes to update title bar only when needed
        if hasattr(self.navigationInterface, 'displayModeChanged'):
            self.navigationInterface.displayModeChanged.connect(self._on_navigation_display_mode_changed)
        
    def _on_navigation_display_mode_changed(self) -> None:
        """Handle navigation display mode change - update title bar position."""
        # Only update when navigation actually expands/collapses, not during page switching
        if not self._is_switching_page:
            self._adjust_title_bar_deferred()

    def resizeEvent(self, event) -> None:
        """
        Handle resize event to adjust TitleBar position dynamically.
        This enables the 'Navigation 2' look where the title bar sits 
        to the right of the navigation sidebar.
        Optimized to reduce unnecessary updates during page switching.
        """
        # Only update title bar position if window is actually resizing
        # Skip updates during page switching to reduce lag
        if (hasattr(self, 'titleBar') and self.titleBar and 
            event.size() != event.oldSize() and 
            not self._is_switching_page):
            self._adjust_title_bar_deferred()
        
        super().resizeEvent(event)
    
    def _adjust_title_bar_deferred(self) -> None:
        """Deferred title bar adjustment to improve performance during page switching."""
        if hasattr(self, 'titleBar') and self.titleBar:
            # Cache navigation width to avoid repeated queries
            nav_width = self.navigationInterface.width()
            # Only update if width actually changed
            if nav_width != self._cached_nav_width:
                self._cached_nav_width = nav_width
                self.titleBar.move(nav_width, 0)
                self.titleBar.resize(self.width() - nav_width, self.titleBar.height())

    def _apply_font_scale(self) -> None:
        """Apply a global font size based on current window size."""
        w = max(1, self.width())
        h = max(1, self.height())
        bw, bh = self._font_base_size
        scale = min(w / max(1, bw), h / max(1, bh))
        scale = max(0.85, min(scale, 1.25))
        pt = int(round(self._font_base_pt * scale))
        pt = max(self._font_min_pt, min(pt, self._font_max_pt))

        app = QApplication.instance()
        if not app:
            return

        font = app.font() or QFont()
        if not font.family():
            font.setFamily("Segoe UI")
        font.setPointSize(pt)
        app.setFont(font)
    
    def _switch_to_page(self, page_id: str) -> None:
        """
        Optimized page switching method.
        Reduces lag by batching updates and avoiding unnecessary repaints.
        Prevents title bar adjustments during page switching when navigation is expanded.
        """
        if page_id in self.task_pages:
            page = self.task_pages[page_id]
            
            # Mark that we're switching pages to prevent unnecessary title bar updates
            self._is_switching_page = True
            
            # Temporarily disable updates to batch paint operations
            self.stackedWidget.setUpdatesEnabled(False)
            try:
                self.stackedWidget.setCurrentWidget(page)
                # Update navigation selection after page switch
                if hasattr(self.navigationInterface, 'setCurrentItem'):
                    try:
                        self.navigationInterface.setCurrentItem(page_id)
                    except Exception:
                        pass
            finally:
                # Re-enable updates and trigger a single repaint
                self.stackedWidget.setUpdatesEnabled(True)
                self.stackedWidget.update()
                self._is_switching_page = False
    
    def _connect_signals(self) -> None:
        """Connect application context signals."""
        self.app_context.login_state_changed.connect(self._on_app_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_app_user_info_changed)
    
    def _init_pages(self) -> None:
        """Initialize all pages (home and task pages)."""
        #region agent log
        try:
            import qfluentwidgets as qfw
            import inspect
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
        home_page.setObjectName("home")
        
        # Set up navigation callback - switch to AP page using objectName
        # Optimized: use cached page reference instead of searching
        # NOTE: These callbacks are part of UI navigation behavior, not business logic.
        def navigate_to_ap():
            if "ap" in self.task_pages:
                self._switch_to_page("ap")
                    
        home_page._navigate_to_ap = navigate_to_ap
        
        # AR navigation
        def navigate_to_ar():
            if "ar" in self.task_pages:
                self._switch_to_page("ar")
        home_page._navigate_to_ar = navigate_to_ar
        home_page._trigger_login = self._on_login_clicked
        
        # Add home page to navigation (first item)
        try:
            # Load theme-aware monochrome icon for Home, fallback to FluentIcon.HOME
            def load_icon(icon_name: str):
                icons_dir = Path(__file__).resolve().parent.parent / "resources" / "icons"
                # Use *_light for dark theme (light icon on dark bg), *_dark for light theme
                suffix = "_light" if isDarkTheme() else "_dark"
                svg_path = icons_dir / f"{icon_name}{suffix}.svg"
                if svg_path.exists():
                    return QIcon(str(svg_path))
                png_path = icons_dir / f"{icon_name}{suffix}.png"
                if png_path.exists():
                    return QIcon(str(png_path))
                # Fallback to base name without suffix
                base_svg = icons_dir / f"{icon_name}.svg"
                if base_svg.exists():
                    return QIcon(str(base_svg))
                base_png = icons_dir / f"{icon_name}.png"
                if base_png.exists():
                    return QIcon(str(base_png))
                return FluentIcon.HOME
            
            home_icon = load_icon("home")
            
            self.addSubInterface(
                interface=home_page,
                icon=home_icon,
                text="Home",
                position=NavigationItemPosition.TOP
            )
        except Exception as e:
            _agent_log("H7", "MainWindow._init_pages", "Failed to add home page", {"error": str(e)})
            raise
        
        # Create task pages with lazy initialization optimization
        # Pages are created immediately but can be optimized for faster switching
        self.task_pages["ap"] = APTaskPage()
        self.task_pages["ar"] = ARTaskPage()
        self.task_pages["project_status"] = ProjectStatusTaskPage()
        self.task_pages["revenue_accrual"] = RevenueAccrualTaskPage()
        self.task_pages["statement_check"] = StatementCheckTaskPage()

        # Add task pages to navigation and stacked widget
        for task_id, page in self.task_pages.items():
            page.setObjectName(task_id)
            # Keep default paint behavior to avoid visual artifacts
            try:
                # Load theme-aware monochrome icons for task pages, with FluentIcon fallback
                def load_icon(icon_name: str):
                    icons_dir = Path(__file__).resolve().parent.parent / "resources" / "icons"
                    suffix = "_light" if isDarkTheme() else "_dark"
                    svg_path = icons_dir / f"{icon_name}{suffix}.svg"
                    if svg_path.exists():
                        return QIcon(str(svg_path))
                    png_path = icons_dir / f"{icon_name}{suffix}.png"
                    if png_path.exists():
                        return QIcon(str(png_path))
                    base_svg = icons_dir / f"{icon_name}.svg"
                    if base_svg.exists():
                        return QIcon(str(base_svg))
                    base_png = icons_dir / f"{icon_name}.png"
                    if base_png.exists():
                        return QIcon(str(base_png))
                    return FluentIcon.DOCUMENT
                
                icon_map = {
                    "ap": load_icon("ap"),
                    "ar": load_icon("ar"),
                    "project_status": load_icon("project_status"),
                    "revenue_accrual": load_icon("revenue_accrual"),
                    "statement_check": load_icon("statement_check"),
                }
                icon = icon_map.get(task_id, FluentIcon.DOCUMENT)
                
                self.addSubInterface(
                    interface=page,
                    icon=icon,
                    text=page.task_name,
                    position=NavigationItemPosition.TOP
                )
            except Exception as exc:
                _agent_log("H7", "MainWindow._init_pages", "addSubInterface error", {"error": str(exc)})
                raise
            
            # Connect signals
            page.task_started.connect(self._on_task_started)
            page.task_finished.connect(self._on_task_finished)
            page.task_failed.connect(self._on_task_failed)
        
        # Add scheduled tasks page
        scheduled_tasks_page = ScheduledTasksPage()
        scheduled_tasks_page.setObjectName("scheduled_tasks")
        try:
            # Load theme-aware monochrome icon for Scheduled Tasks, fallback to FluentIcon.CALENDAR
            def load_icon(icon_name: str):
                icons_dir = Path(__file__).resolve().parent.parent / "resources" / "icons"
                suffix = "_light" if isDarkTheme() else "_dark"
                svg_path = icons_dir / f"{icon_name}{suffix}.svg"
                if svg_path.exists():
                    return QIcon(str(svg_path))
                png_path = icons_dir / f"{icon_name}{suffix}.png"
                if png_path.exists():
                    return QIcon(str(png_path))
                base_svg = icons_dir / f"{icon_name}.svg"
                if base_svg.exists():
                    return QIcon(str(base_svg))
                base_png = icons_dir / f"{icon_name}.png"
                if base_png.exists():
                    return QIcon(str(base_png))
                return FluentIcon.CALENDAR
            
            schedule_icon = load_icon("schedule")

            self.addSubInterface(
                interface=scheduled_tasks_page,
                icon=schedule_icon,
                text="Scheduled Tasks",
                position=NavigationItemPosition.TOP
            )
        except Exception as e:
            try:
                self.addSubInterface(
                    interface=scheduled_tasks_page,
                    icon=FluentIcon.CALENDAR,
                    text="Scheduled Tasks",
                    position=NavigationItemPosition.TOP
                )
            except Exception:
                pass
        
        # Setup bottom navigation (avatar and settings)
        self._setup_bottom_navigation()
    
    @Slot(str)
    def _on_task_started(self, task_id: str) -> None:
        """Handle task started signal."""
        for page in self.task_pages.values():
            page._set_execution_enabled(False)
        
        if task_id in self.task_pages:
            self.current_worker = self.task_pages[task_id].current_worker
        
        original_title = "QCA Accounting Automation Tool"
        self.setWindowTitle(f"{original_title} - Executing: {task_id}")
    
    @Slot(str, dict)
    def _on_task_finished(self, task_id: str, result: dict) -> None:
        """Handle worker finished signal."""
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        self.current_worker = None
        self.setWindowTitle("QCA Accounting Automation Tool")
        
        # Show completion dialog when task finished successfully
        try:
            success = bool(result.get("success"))
            message = result.get("message", "Task completed.")
            if success:
                dialog = MessageBox(
                    title="Task Completed",
                    content=message,
                    parent=self,
                )
                dialog.exec()
        except Exception:
            # Silently ignore errors from completion dialog
            pass
    
    @Slot(str, str)
    def _on_task_failed(self, task_id: str, error: str) -> None:
        """Handle worker failed signal."""
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        self.current_worker = None
        self.setWindowTitle("QCA Accounting Automation Tool")
    
    def _check_login_status(self) -> None:
        """Check login status using background worker."""
        if self.check_login_worker and self.check_login_worker.isRunning():
            return
        self.check_login_worker = CheckLoginWorker()
        self.check_login_worker.finished.connect(self._on_check_login_finished)
        self.check_login_worker.start()
    
    @Slot(dict)
    def _on_check_login_finished(self, result: Dict[str, Any]) -> None:
        """Handle check login finished signal."""
        is_logged_in = result.get("is_logged_in", False)
        user_info = {}
        if is_logged_in:
            user_info = result.get("user_info", {})
        self.app_context.set_login_state(is_logged_in, user_info if user_info else None)
        self._update_login_ui()
    
    @Slot()
    def _on_login_clicked(self) -> None:
        """Handle login button click - show login dialog."""
        if self.login_worker and self.login_worker.isRunning():
            return
        self.login_dialog = LoginDialog(self)
        self.login_dialog.login_requested.connect(self._start_login)
        self.login_dialog.exec()
    
    def _start_login(self) -> None:
        """Start login process."""
        if self.login_worker and self.login_worker.isRunning():
            return
        self.login_worker = LoginWorker(
            headless=False,
            log_callback=self._on_login_log_received
        )
        self.login_worker.log_signal.connect(self._on_login_log_received)
        self.login_worker.finished.connect(self._on_login_finished)
        if self.login_dialog:
            self.login_dialog._set_loading_state(True, "Opening browser for authentication...")
        self.login_worker.start()
    
    @Slot(str, str)
    def _on_login_log_received(self, level: str, message: str) -> None:
        """Handle log signal from login worker."""
        if not self.login_dialog:
            return
        if level == "INFO":
            if "Starting" in message or "Opening" in message:
                status_msg = "Opening browser..."
                self.login_dialog._set_loading_state(True, status_msg)
            elif "completed" in message.lower() or "Verifying" in message:
                status_msg = "Verifying login..."
                self.login_dialog._set_loading_state(True, status_msg)
            else:
                self.login_dialog._set_loading_state(True, message)
        elif level == "SUCCESS":
            self.login_dialog.show_success()
        elif level == "ERROR":
            self.login_dialog.show_error(message)
    
    @Slot(dict)
    def _on_login_finished(self, result: Dict[str, Any]) -> None:
        """Handle login finished signal."""
        success = result.get("success", False)
        is_logged_in = result.get("is_logged_in", False)
        user_info = {}
        if success and is_logged_in:
            user_info = result.get("user_info", {})
        self.app_context.set_login_state(is_logged_in, user_info if user_info else None)
        self._update_login_ui()
        if success and is_logged_in:
            self._check_login_status()
    
    def _update_login_ui(self) -> None:
        """Update login UI based on app context state."""
        self._update_avatar_widget()
    
    @Slot(bool)
    def _on_app_login_state_changed(self, is_logged_in: bool) -> None:
        """Handle app context login state changed signal."""
        self._update_login_ui()
    
    @Slot(dict)
    def _on_app_user_info_changed(self, user_info: dict) -> None:
        """Handle app context user info changed signal."""
        self._update_login_ui()
    
    def _setup_bottom_navigation(self) -> None:
        """Setup bottom navigation with AvatarWidget and Settings."""
        if not self.navigationInterface:
            return
        
        # Create custom avatar widget
        self.avatar_widget = CustomAvatarWidget(self)
        
        # Add avatar widget to bottom navigation
        self.navigationInterface.addWidget(
            routeKey='avatar',
            widget=self.avatar_widget,
            onClick=self._on_avatar_clicked,
            position=NavigationItemPosition.BOTTOM
        )
        
        # Create empty settings widget
        self.settings_widget = SettingsWidget(self)
        self.settings_widget.setObjectName("settings")
        
        self.stackedWidget.addWidget(self.settings_widget)
        
        self.addSubInterface(
            interface=self.settings_widget,
            icon=FIF.SETTING,
            text='Settings',
            position=NavigationItemPosition.BOTTOM
        )
    
    def _on_avatar_clicked(self) -> None:
        """Handle avatar widget click event"""
        import tools.util as util
        user_email = util.check_login()
        if not user_email:
            self._on_login_clicked()
    
    def _update_avatar_widget(self) -> None:
        """Update avatar widget to reflect current login state."""
        if not self.avatar_widget:
            return
        self.avatar_widget._load_avatar()
        self.avatar_widget.update()
    
    def _apply_stacked_widget_style(self) -> None:
        """Apply theme-aware background to stacked widget."""
        try:
            if not hasattr(self, 'stackedWidget') or not self.stackedWidget:
                return
            bg = ThemeColors.background_primary()
            text = ThemeColors.text_primary()
            self.stackedWidget.setStyleSheet(f"""
                QStackedWidget {{
                    background-color: {bg};
                    color: {text};
                }}
            """)
            # Set palette for proper color inheritance
            palette = self.stackedWidget.palette()
            palette.setColor(QPalette.ColorRole.WindowText, QColor(text))
            palette.setColor(QPalette.ColorRole.Window, QColor(bg))
            palette.setColor(QPalette.ColorRole.Base, QColor(bg))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(bg))
            self.stackedWidget.setPalette(palette)
        except Exception:
            # Silently ignore errors to prevent crashes during initialization
            pass
    
    def _update_all_pages_theme(self) -> None:
        """Update theme styles for all pages."""
        # Check if initialization is complete
        if not hasattr(self, '_initialization_complete') or not self._initialization_complete:
            return
        
        if not hasattr(self, 'task_pages'):
            return
        
        try:
            # Update stacked widget style
            if hasattr(self, 'stackedWidget') and self.stackedWidget:
                self._apply_stacked_widget_style()
            
            # Update all task pages
            if hasattr(self, 'task_pages') and self.task_pages:
                for page in self.task_pages.values():
                    try:
                        if hasattr(page, '_apply_page_background'):
                            page._apply_page_background()
                        if hasattr(page, '_apply_scroll_area_style'):
                            page._apply_scroll_area_style()
                        if hasattr(page, '_apply_content_widget_style'):
                            page._apply_content_widget_style()
                        if hasattr(page, '_apply_log_viewer_style'):
                            page._apply_log_viewer_style()
                    except Exception:
                        # Silently ignore errors for individual pages
                        pass
            
            # Update home page if it exists
            if hasattr(self, 'stackedWidget') and self.stackedWidget.count() > 0:
                try:
                    home_page = self.stackedWidget.widget(0)
                    if home_page and hasattr(home_page, '_apply_theme_styles'):
                        home_page._apply_theme_styles()
                except Exception:
                    # Silently ignore errors for home page
                    pass
        except Exception:
            # Silently ignore errors during theme update to prevent crashes
            pass
    
    def changeEvent(self, event: QEvent) -> None:
        """
        Handle change events, including theme changes.
        
        Args:
            event: Change event
        """
        super().changeEvent(event)
        # Temporarily disable automatic theme updates in changeEvent to prevent crashes
        # Theme updates will be handled explicitly when user switches theme
        # This prevents crashes during initialization
        pass
    
    def update_theme_for_all_pages(self) -> None:
        """
        Public method to update theme for all pages.
        Called explicitly when user switches theme.
        """
        try:
            if hasattr(self, '_initialization_complete') and self._initialization_complete:
                # Update stacked widget style
                self._apply_stacked_widget_style()
                
                # Update current visible page
                current_widget = self.stackedWidget.currentWidget()
                if current_widget:
                    # Trigger showEvent to apply theme styles
                    if hasattr(current_widget, 'showEvent'):
                        try:
                            current_widget.showEvent(QShowEvent())
                        except Exception:
                            pass
                
                # Update all task pages (they will update when shown)
                if hasattr(self, 'task_pages') and self.task_pages:
                    for page in self.task_pages.values():
                        try:
                            # Only update if page is visible
                            if page.isVisible():
                                page.showEvent(QShowEvent())
                        except Exception:
                            pass

                # Also update navigation icons to match new theme
                try:
                    nav = getattr(self, 'navigationInterface', None)
                    if nav and hasattr(nav, 'panel') and hasattr(nav.panel, 'widget'):
                        from qfluentwidgets.components.navigation import NavigationPushButton

                        def _load_nav_icon(icon_name: str, default_icon: FluentIcon) -> QIcon:
                            icons_dir = Path(__file__).resolve().parent.parent / "resources" / "icons"
                            suffix = "_light" if isDarkTheme() else "_dark"
                            svg_path = icons_dir / f"{icon_name}{suffix}.svg"
                            if svg_path.exists():
                                return QIcon(str(svg_path))
                            png_path = icons_dir / f"{icon_name}{suffix}.png"
                            if png_path.exists():
                                return QIcon(str(png_path))
                            base_svg = icons_dir / f"{icon_name}.svg"
                            if base_svg.exists():
                                return QIcon(str(base_svg))
                            base_png = icons_dir / f"{icon_name}.png"
                            if base_png.exists():
                                return QIcon(str(base_png))
                            return default_icon

                        # Home icon
                        try:
                            w = nav.panel.widget('home')
                            if isinstance(w, NavigationPushButton):
                                w.setIcon(_load_nav_icon('home', FluentIcon.HOME))
                        except Exception:
                            pass

                        # Task icons
                        for route_key, icon_name, default_icon in [
                            ('ap', 'ap', FluentIcon.DOCUMENT),
                            ('ar', 'ar', FluentIcon.DOCUMENT),
                            ('project_status', 'project_status', FluentIcon.DOCUMENT),
                            ('revenue_accrual', 'revenue_accrual', FluentIcon.DOCUMENT),
                            ('statement_check', 'statement_check', FluentIcon.DOCUMENT),
                        ]:
                            try:
                                w = nav.panel.widget(route_key)
                                if isinstance(w, NavigationPushButton):
                                    w.setIcon(_load_nav_icon(icon_name, default_icon))
                            except Exception:
                                pass

                        # Scheduled tasks icon
                        try:
                            w = nav.panel.widget('scheduled_tasks')
                            if isinstance(w, NavigationPushButton):
                                w.setIcon(_load_nav_icon('schedule', FluentIcon.CALENDAR))
                        except Exception:
                            pass
                except Exception:
                    # Silently ignore navigation icon update errors
                    pass
        except Exception:
            # Silently ignore all errors
            pass