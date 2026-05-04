"""
Main window for QCA Accounting Automation Tool.
"""
from typing import Dict, Optional, Any
from pathlib import Path
import ctypes
from ctypes import wintypes

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QFrame,
)
from PySide6.QtCore import Slot, Qt, QRect, QEvent
from PySide6.QtGui import QFont, QIcon, QPainter, QBrush, QColor, QPalette, QShowEvent

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

from ui.pages.home_page import HomePage
from ui.pages.ap_task_page import APTaskPage
from ui.pages.ar_task_page import ARTaskPage
from ui.pages.project_status_task_page import ProjectStatusTaskPage
from ui.pages.revenue_accrual_task_page import RevenueAccrualTaskPage
from ui.pages.statement_check_task_page import StatementCheckTaskPage
from ui.pages.report_export_task_page import ReportExportTaskPage
from ui.pages.base_task_page import BaseTaskPage
from ui.pages.scheduled_tasks_page import ScheduledTasksPage
from ui.services.app_context import get_app_context
from ui.widgets.login_dialog import LoginDialog
from ui.widgets.settings_page import SettingsPage
from ui.utils.theme_colors import ThemeColors
from workers.base_worker import BaseWorker
from workers.check_login_worker import CheckLoginWorker
from workers.login_worker import LoginWorker


# Windows API structures and flags for taskbar flashing
class FLASHWINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwnd", wintypes.HWND),
        ("dwFlags", wintypes.DWORD),
        ("uCount", wintypes.UINT),
        ("dwTimeout", wintypes.DWORD),
    ]


FLASHW_CAPTION = 0x00000001
FLASHW_TRAY = 0x00000002
FLASHW_ALL = FLASHW_CAPTION | FLASHW_TRAY
FLASHW_TIMERNOFG = 0x0000000C


def flash_window(widget: QWidget, count: int = 5):
    hwnd = int(widget.winId())
    fInfo = FLASHWINFO(
        cbSize=ctypes.sizeof(FLASHWINFO),
        hwnd=hwnd,
        dwFlags=FLASHW_ALL | FLASHW_TIMERNOFG,
        uCount=count,
        dwTimeout=0,
    )
    ctypes.windll.user32.FlashWindowEx(ctypes.byref(fInfo))


class CustomAvatarWidget(NavigationWidget):
    """Avatar widget pinned to the bottom of the navigation sidebar."""

    def __init__(self, parent=None):
        super().__init__(isSelectable=False, parent=parent)
        self.initials = ""
        self.user_email = None

        self.app_context = get_app_context()
        self.app_context.login_state_changed.connect(self._on_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_user_info_changed)
        self._load_avatar()

    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        self._load_avatar()
        self.update()

    def _on_user_info_changed(self, user_info: dict) -> None:
        self._load_avatar()
        self.update()

    def _load_avatar(self) -> None:
        """Populate initials from AppContext, falling back to a direct login check."""
        user_email = self.app_context.user_email
        if not user_email:
            # AppContext may not have updated yet; check directly and sync it.
            import tools.util as util
            user_email = util.check_login()
            if user_email:
                self.app_context.set_login_state(True, {"EMail": user_email, "email": user_email})

        self.user_email = user_email
        if user_email:
            prefix = user_email.split("@")[0]
            parts = prefix.split(".")
            if len(parts) >= 2:
                self.initials = (parts[0][0] + parts[1][0]).upper()
            else:
                self.initials = prefix[:2].upper() if len(prefix) >= 2 else prefix[0].upper()
        else:
            self.initials = "?"

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(QPainter.SmoothPixmapTransform | QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)

        if self.isPressed:
            painter.setOpacity(0.7)

        if self.isEnter:
            c = 255 if isDarkTheme() else 0
            painter.setBrush(QColor(c, c, c, 10))
            painter.drawRoundedRect(self.rect(), 5, 5)

        avatar_rect = QRect(12, 6, 24, 24) if self.isCompacted else QRect(8, 6, 24, 24)
        painter.setBrush(
            QBrush(QColor(100, 150, 200) if not isDarkTheme() else QColor(70, 120, 170))
        )
        painter.drawEllipse(avatar_rect)

        painter.setPen(Qt.white)
        font = QFont("Segoe UI")
        font.setPixelSize(12)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(avatar_rect, Qt.AlignCenter, self.initials)

        # Draw the email prefix label only in the expanded (non-compacted) state.
        if not self.isCompacted and self.user_email:
            painter.setPen(Qt.white if isDarkTheme() else Qt.black)
            font = QFont("Segoe UI")
            font.setPixelSize(14)
            painter.setFont(font)
            painter.drawText(QRect(44, 0, 255, 36), Qt.AlignVCenter, self.user_email.split("@")[0])


class CustomTitleBar(TitleBar):
    """
    Title bar with app icon and title on the left.
    Adapted from PyQt-Fluent-Widgets navigation2/demo.py.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.iconLabel = QLabel(self)
        self.titleLabel = QLabel(self)
        self.iconLabel.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.titleLabel.setAttribute(Qt.WA_TransparentForMouseEvents)

        self.titleLabel.setObjectName("titleLabel")
        self.titleLabel.setFont(QFont("Segoe UI", 12))
        self.iconLabel.setFixedSize(18, 18)

        self.hBoxLayout.insertSpacing(0, 20)
        self.hBoxLayout.insertWidget(1, self.iconLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.hBoxLayout.insertWidget(2, self.titleLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)

        self.window().windowIconChanged.connect(self.setIcon)
        self.window().windowTitleChanged.connect(self.setTitle)

    def setTitle(self, title: str) -> None:
        self.titleLabel.setText(title)
        self.titleLabel.adjustSize()

    def setIcon(self, icon) -> None:
        self.iconLabel.setPixmap(QIcon(icon).pixmap(18, 18))


class SettingsWidget(QFrame):
    """Wrapper frame that hosts the SettingsPage."""

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setObjectName("settings")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.settings_page = SettingsPage(self)
        layout.addWidget(self.settings_page)


class MainWindow(FluentWindow):
    """
    Main application window.
    Left panel: navigation sidebar.  Right panel: task content area.
    """

    def __init__(self, startup_progress=None):
        self._startup_progress = startup_progress
        super().__init__()

        self._report_progress(40, "Initializing window...")
        self._init_window()

        self._report_progress(50, "Loading application context...")
        self.app_context = get_app_context()
        self.current_worker: Optional[BaseWorker] = None
        self.task_pages: Dict[str, BaseTaskPage] = {}
        self._initialization_complete = False

        self._report_progress(60, "Setting up navigation...")
        self._setup_ui()
        self._report_progress(70, "Initializing pages...")
        self._init_pages()
        self._report_progress(80, "Connecting signals...")
        self._connect_signals()
        self._report_progress(85, "Checking login status...")
        self._check_login_status()

        self._initialization_complete = True
        self._report_progress(90, "Ready")

    def _report_progress(self, progress: int, message: str) -> None:
        if self._startup_progress:
            try:
                self._startup_progress(progress, message)
            except Exception:
                pass

    def center_window(self) -> None:
        frame = self.frameGeometry()
        frame.moveCenter(QApplication.primaryScreen().availableGeometry().center())
        self.move(frame.topLeft())

    def _init_window(self) -> None:
        """Set up title bar, window dimensions, and initial theme."""
        self.setTitleBar(CustomTitleBar(self))
        self.setWindowTitle("QCA Accounting Automation Tool")
        self.resize(1200, 800)
        self.setMinimumSize(900, 650)
        self.center_window()
        try:
            setTheme(Theme.LIGHT)
        except Exception:
            pass

    def _setup_ui(self) -> None:
        """Configure navigation interface settings and shared UI state."""
        self.navigationInterface.setExpandWidth(280)
        self.navigationInterface.setMinimumExpandWidth(300)
        self.stackedWidget.setContentsMargins(0, 0, 0, 0)

        try:
            self._apply_stacked_widget_style()
        except Exception:
            pass

        self.check_login_worker: Optional[CheckLoginWorker] = None
        self.login_worker: Optional[LoginWorker] = None
        self.login_dialog: Optional[LoginDialog] = None
        self.avatar_widget: Optional[CustomAvatarWidget] = None

        # Track cached nav width to avoid redundant title bar moves.
        self._cached_nav_width = 0
        self._is_switching_page = False

        # Expand the sidebar by default on startup.
        self.navigationInterface.expand(useAni=False)

        if hasattr(self.navigationInterface, "displayModeChanged"):
            self.navigationInterface.displayModeChanged.connect(
                self._on_navigation_display_mode_changed
            )

    def _on_navigation_display_mode_changed(self) -> None:
        """Update the title bar position when the nav panel expands or collapses."""
        if not self._is_switching_page:
            self._adjust_title_bar_deferred()

    def resizeEvent(self, event) -> None:
        if (
            hasattr(self, "titleBar")
            and self.titleBar
            and event.size() != event.oldSize()
            and not self._is_switching_page
        ):
            self._adjust_title_bar_deferred()
        super().resizeEvent(event)

    def _adjust_title_bar_deferred(self) -> None:
        """Shift the title bar right to sit alongside the navigation panel."""
        if not (hasattr(self, "titleBar") and self.titleBar):
            return
        nav_width = self.navigationInterface.width()
        if nav_width != self._cached_nav_width:
            self._cached_nav_width = nav_width
            self.titleBar.move(nav_width, 0)
            self.titleBar.resize(self.width() - nav_width, self.titleBar.height())

    def _switch_to_page(self, page_id: str) -> None:
        """Switch to a task page while suppressing redundant title bar redraws."""
        if page_id not in self.task_pages:
            return
        self._is_switching_page = True
        self.stackedWidget.setUpdatesEnabled(False)
        try:
            self.stackedWidget.setCurrentWidget(self.task_pages[page_id])
            if hasattr(self.navigationInterface, "setCurrentItem"):
                try:
                    self.navigationInterface.setCurrentItem(page_id)
                except Exception:
                    pass
        finally:
            self.stackedWidget.setUpdatesEnabled(True)
            self.stackedWidget.update()
            self._is_switching_page = False

    def _connect_signals(self) -> None:
        self.app_context.login_state_changed.connect(self._on_app_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_app_user_info_changed)

    def _load_icon(self, icon_name: str, default: FluentIcon, suffix: str) -> QIcon:
        """
        Load a themed SVG/PNG icon from resources/icons, trying the suffixed variant
        first, then the unsuffixed base name, and finally returning *default*.
        """
        icons_dir = Path(__file__).resolve().parent.parent / "resources" / "icons"
        for ext in ("svg", "png"):
            p = icons_dir / f"{icon_name}{suffix}.{ext}"
            if p.exists():
                return QIcon(str(p))
        for ext in ("svg", "png"):
            p = icons_dir / f"{icon_name}.{ext}"
            if p.exists():
                return QIcon(str(p))
        return default

    def _init_pages(self) -> None:
        """Create and register all navigation pages."""
        home_page = HomePage()
        home_page.setObjectName("home")

        def navigate_to_ap():
            if "ap" in self.task_pages:
                self._switch_to_page("ap")

        def navigate_to_ar():
            if "ar" in self.task_pages:
                self._switch_to_page("ar")

        home_page._navigate_to_ap = navigate_to_ap
        home_page._navigate_to_ar = navigate_to_ar
        home_page._trigger_login = self._on_login_clicked

        # Home and schedule icons use light variants on dark backgrounds.
        contrast_suffix = "_light" if isDarkTheme() else "_dark"

        self.addSubInterface(
            interface=home_page,
            icon=self._load_icon("home", FluentIcon.HOME, contrast_suffix),
            text="Home",
            position=NavigationItemPosition.TOP,
        )

        self.task_pages["ap"] = APTaskPage()
        self.task_pages["ar"] = ARTaskPage()
        self.task_pages["project_status"] = ProjectStatusTaskPage()
        self.task_pages["revenue_accrual"] = RevenueAccrualTaskPage()
        self.task_pages["statement_check"] = StatementCheckTaskPage()
        self.task_pages["report_export"] = ReportExportTaskPage()

        # Task page icons use the opposite suffix convention from home/schedule.
        task_suffix = "_dark" if isDarkTheme() else "_light"
        icon_map = {
            task_id: self._load_icon(task_id, FluentIcon.DOCUMENT, task_suffix)
            for task_id in self.task_pages
        }

        for task_id, page in self.task_pages.items():
            page.setObjectName(task_id)
            self.addSubInterface(
                interface=page,
                icon=icon_map[task_id],
                text=page.task_name,
                position=NavigationItemPosition.TOP,
            )
            page.task_started.connect(self._on_task_started)
            page.task_finished.connect(self._on_task_finished)
            page.task_failed.connect(self._on_task_failed)

        scheduled_tasks_page = ScheduledTasksPage()
        scheduled_tasks_page.setObjectName("scheduled_tasks")
        try:
            self.addSubInterface(
                interface=scheduled_tasks_page,
                icon=self._load_icon("schedule", FluentIcon.CALENDAR, contrast_suffix),
                text="Scheduled Tasks",
                position=NavigationItemPosition.TOP,
            )
        except Exception:
            # Fall back to a built-in FluentIcon if the custom icon cannot be loaded.
            try:
                self.addSubInterface(
                    interface=scheduled_tasks_page,
                    icon=FluentIcon.CALENDAR,
                    text="Scheduled Tasks",
                    position=NavigationItemPosition.TOP,
                )
            except Exception:
                pass

        self._setup_bottom_navigation()

    @Slot(str)
    def _on_task_started(self, task_id: str) -> None:
        for page in self.task_pages.values():
            page._set_execution_enabled(False)
        if task_id in self.task_pages:
            self.current_worker = self.task_pages[task_id].current_worker
        self.setWindowTitle(f"QCA Accounting Automation Tool - Executing: {task_id}")

    @Slot(str, dict)
    def _on_task_finished(self, task_id: str, result: dict) -> None:
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        self.current_worker = None
        self.setWindowTitle("QCA Accounting Automation Tool")
        flash_window(self, count=10)
        try:
            if result.get("success"):
                MessageBox(
                    title="Task Completed",
                    content=result.get("message", "Task completed."),
                    parent=self,
                ).exec()
        except Exception:
            pass

    @Slot(str, str)
    def _on_task_failed(self, task_id: str, error: str) -> None:
        for page in self.task_pages.values():
            page._set_execution_enabled(True)
        self.current_worker = None
        self.setWindowTitle("QCA Accounting Automation Tool")

    def _check_login_status(self) -> None:
        if self.check_login_worker and self.check_login_worker.isRunning():
            return
        self.check_login_worker = CheckLoginWorker()
        self.check_login_worker.finished.connect(self._on_check_login_finished)
        self.check_login_worker.start()

    @Slot(dict)
    def _on_check_login_finished(self, result: Dict[str, Any]) -> None:
        is_logged_in = result.get("is_logged_in", False)
        user_info = result.get("user_info", {}) if is_logged_in else {}
        self.app_context.set_login_state(is_logged_in, user_info or None)
        self._update_login_ui()

    @Slot()
    def _on_login_clicked(self) -> None:
        if self.login_worker and self.login_worker.isRunning():
            return
        self.login_dialog = LoginDialog(self)
        self.login_dialog.login_requested.connect(self._start_login)
        self.login_dialog.exec()

    def _start_login(self) -> None:
        if self.login_worker and self.login_worker.isRunning():
            return
        self.login_worker = LoginWorker(
            headless=False, log_callback=self._on_login_log_received
        )
        self.login_worker.log_signal.connect(self._on_login_log_received)
        self.login_worker.finished.connect(self._on_login_finished)
        if self.login_dialog:
            self.login_dialog._set_loading_state(True, "Opening browser for authentication...")
        self.login_worker.start()

    @Slot(str, str)
    def _on_login_log_received(self, level: str, message: str) -> None:
        if not self.login_dialog:
            return
        if level == "INFO":
            if "Starting" in message or "Opening" in message:
                self.login_dialog._set_loading_state(True, "Opening browser...")
            elif "completed" in message.lower() or "Verifying" in message:
                self.login_dialog._set_loading_state(True, "Verifying login...")
            else:
                self.login_dialog._set_loading_state(True, message)
        elif level == "SUCCESS":
            self.login_dialog.show_success()
        elif level == "ERROR":
            self.login_dialog.show_error(message)

    @Slot(dict)
    def _on_login_finished(self, result: Dict[str, Any]) -> None:
        success = result.get("success", False)
        is_logged_in = result.get("is_logged_in", False)
        user_info = result.get("user_info", {}) if (success and is_logged_in) else {}
        self.app_context.set_login_state(is_logged_in, user_info or None)
        self._update_login_ui()
        if success and is_logged_in:
            self._check_login_status()

    def _update_login_ui(self) -> None:
        self._update_avatar_widget()

    @Slot(bool)
    def _on_app_login_state_changed(self, is_logged_in: bool) -> None:
        self._update_login_ui()

    @Slot(dict)
    def _on_app_user_info_changed(self, user_info: dict) -> None:
        self._update_login_ui()

    def _setup_bottom_navigation(self) -> None:
        if not self.navigationInterface:
            return

        self.avatar_widget = CustomAvatarWidget(self)
        self.navigationInterface.addWidget(
            routeKey="avatar",
            widget=self.avatar_widget,
            onClick=self._on_avatar_clicked,
            position=NavigationItemPosition.BOTTOM,
        )

        self.settings_widget = SettingsWidget(self)
        self.settings_widget.setObjectName("settings")
        self.stackedWidget.addWidget(self.settings_widget)
        self.addSubInterface(
            interface=self.settings_widget,
            icon=FIF.SETTING,
            text="Settings",
            position=NavigationItemPosition.BOTTOM,
        )

    def _on_avatar_clicked(self) -> None:
        import tools.util as util
        if not util.check_login():
            self._on_login_clicked()

    def _update_avatar_widget(self) -> None:
        if self.avatar_widget:
            self.avatar_widget._load_avatar()
            self.avatar_widget.update()

    def _apply_stacked_widget_style(self) -> None:
        """Apply theme-aware background and palette to the stacked widget."""
        try:
            if not hasattr(self, "stackedWidget") or not self.stackedWidget:
                return
            bg = ThemeColors.background_primary()
            text = ThemeColors.text_primary()
            self.stackedWidget.setStyleSheet(
                f"QStackedWidget {{ background-color: {bg}; color: {text}; }}"
            )
            palette = self.stackedWidget.palette()
            palette.setColor(QPalette.ColorRole.WindowText, QColor(text))
            palette.setColor(QPalette.ColorRole.Window, QColor(bg))
            palette.setColor(QPalette.ColorRole.Base, QColor(bg))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(bg))
            self.stackedWidget.setPalette(palette)
        except Exception:
            pass

    def changeEvent(self, event: QEvent) -> None:
        # Theme updates are triggered explicitly via update_theme_for_all_pages()
        # rather than here, to avoid crashes during early initialization.
        super().changeEvent(event)

    def update_theme_for_all_pages(self) -> None:
        """Apply the current theme to all pages. Called explicitly on theme change."""
        if not getattr(self, "_initialization_complete", False):
            return
        try:
            self._apply_stacked_widget_style()

            # Trigger showEvent on the currently visible page to refresh its styles.
            current_widget = self.stackedWidget.currentWidget()
            if current_widget and hasattr(current_widget, "showEvent"):
                try:
                    current_widget.showEvent(QShowEvent())
                except Exception:
                    pass

            for page in self.task_pages.values():
                try:
                    if page.isVisible():
                        page.showEvent(QShowEvent())
                except Exception:
                    pass

            nav = getattr(self, "navigationInterface", None)
            if not (nav and hasattr(nav, "panel") and hasattr(nav.panel, "widget")):
                return

            from qfluentwidgets.components.navigation import NavigationPushButton

            # All nav icon updates use the same suffix convention during theme refresh.
            suffix = "_dark" if isDarkTheme() else "_light"
            for route_key, icon_name, default_icon in [
                ("home", "home", FluentIcon.HOME),
                ("ap", "ap", FluentIcon.DOCUMENT),
                ("ar", "ar", FluentIcon.DOCUMENT),
                ("project_status", "project_status", FluentIcon.DOCUMENT),
                ("revenue_accrual", "revenue_accrual", FluentIcon.DOCUMENT),
                ("statement_check", "statement_check", FluentIcon.DOCUMENT),
                ("scheduled_tasks", "schedule", FluentIcon.CALENDAR),
            ]:
                try:
                    w = nav.panel.widget(route_key)
                    if isinstance(w, NavigationPushButton):
                        w.setIcon(self._load_icon(icon_name, default_icon, suffix))
                except Exception:
                    pass
        except Exception:
            pass
