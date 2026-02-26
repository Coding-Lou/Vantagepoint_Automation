"""
Home page (Dashboard) for the application.

Displays overview, login status, recent tasks, and quick access buttons.
"""
from typing import Optional, Any
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PySide6.QtCore import Qt, QTimer

from qfluentwidgets import (
    CardWidget,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
    BodyLabel,
    TitleLabel,
    CaptionLabel,
    MessageBox,
    setTheme,
    Theme,
    isDarkTheme,
)

from ui.services.app_context import get_app_context
from ui.utils.theme_colors import ThemeColors


class HomePage(QWidget):
    """
    Home page with dashboard layout.
    
    Features:
    - Application overview
    - Login status display
    - Quick access buttons
    - Recent tasks (placeholder)
    """
    
    def __init__(self, parent=None):
        """Initialize home page."""
        super().__init__(parent)
        self.app_context = get_app_context()
        self.check_update_worker: Optional[Any] = None
        self._setup_ui()
        self._connect_signals()
        self._update_login_status()
        # Check for updates in background after a short delay
        QTimer.singleShot(2000, self._check_for_updates_background)
    
    def _setup_ui(self) -> None:
        """Setup UI layout."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)
        
        # Title section with update button and theme toggle in top-right
        title_layout = QHBoxLayout()
        title_label = TitleLabel("QCA Accounting Automation Tool")
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        # Update button in top-right corner
        self.update_btn = PushButton("Check for Updates", self, FluentIcon.SYNC)
        self.update_btn.setMinimumHeight(36)
        self.update_btn.clicked.connect(self._on_check_update_clicked)
        title_layout.addWidget(self.update_btn)
        
        # Theme toggle button
        self.theme_btn = PushButton(self._get_theme_button_text(), self, FluentIcon.BRUSH)
        self.theme_btn.setMinimumHeight(36)
        self.theme_btn.setToolTip(self._get_theme_tooltip())
        self.theme_btn.clicked.connect(self._on_theme_toggle_clicked)
        title_layout.addWidget(self.theme_btn)
        
        layout.addLayout(title_layout)
        
        # Description
        desc_label = BodyLabel(
            "Automate accounting tasks including AP remittance, AR statements, "
            "project status reports, and more."
        )
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        # Status cards row
        status_layout = QHBoxLayout()
        status_layout.setSpacing(12)
        
        # Login status card
        self.login_status_card = self._create_status_card(
            "Login Status",
            "Not logged in",
            FluentIcon.PEOPLE,
            "#f44336"  # Red for not logged in
        )
        status_layout.addWidget(self.login_status_card)
        
        # System status card (placeholder)
        system_status_card = self._create_status_card(
            "System Status",
            "Ready",
            FluentIcon.SETTING,
            "#4caf50"  # Green
        )
        status_layout.addWidget(system_status_card)
        
        status_layout.addStretch()
        layout.addLayout(status_layout)
        
        # Quick actions section
        actions_label = TitleLabel("Quick Actions")
        actions_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(actions_label)
        
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(12)
        
        # AP quick action button
        ap_btn = PrimaryPushButton("AP Remittance", self, FluentIcon.DOCUMENT)
        ap_btn.setMinimumHeight(40)
        ap_btn.clicked.connect(self._on_ap_clicked)
        actions_layout.addWidget(ap_btn)
        
        # AR quick action button
        ar_btn = PrimaryPushButton("AR Notification", self, FluentIcon.DOCUMENT)
        ar_btn.setMinimumHeight(40)
        ar_btn.clicked.connect(self._on_ar_clicked)
        actions_layout.addWidget(ar_btn)
        
        # Login button
        self.login_btn = PrimaryPushButton("Sign In", self, FluentIcon.PEOPLE)
        self.login_btn.setMinimumHeight(40)
        self.login_btn.clicked.connect(self._on_login_clicked)
        actions_layout.addWidget(self.login_btn)
        
        actions_layout.addStretch()
        layout.addLayout(actions_layout)
        
        # Recent tasks section (placeholder)
        recent_label = TitleLabel("Recent Tasks")
        recent_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(recent_label)
        
        recent_card = CardWidget(self)
        recent_layout = QVBoxLayout(recent_card)
        recent_layout.setContentsMargins(16, 16, 16, 16)
        
        no_tasks_label = CaptionLabel("No recent tasks")
        no_tasks_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        recent_layout.addWidget(no_tasks_label)
        
        layout.addWidget(recent_card)
        
        layout.addStretch()
    
    def _create_status_card(
        self,
        title: str,
        status: str,
        icon: FluentIcon,
        color: str
    ) -> CardWidget:
        """
        Create a status card widget.
        
        Args:
            title: Card title
            status: Status text
            icon: Icon to display
            color: Status color
            
        Returns:
            CardWidget instance
        """
        card = CardWidget(self)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(8)
        
        # Title
        title_label = BodyLabel(title)
        title_label.setStyleSheet("font-weight: 600;")
        card_layout.addWidget(title_label)
        
        # Status
        status_label = BodyLabel(status)
        status_label.setStyleSheet(f"color: {color}; font-size: 14px;")
        card_layout.addWidget(status_label)
        
        return card
    
    def _connect_signals(self) -> None:
        """Connect application context signals."""
        self.app_context.login_state_changed.connect(self._on_login_state_changed)
        self.app_context.user_info_changed.connect(self._on_user_info_changed)
    
    def _update_login_status(self) -> None:
        """Update login status display."""
        is_logged_in = self.app_context.is_logged_in
        user_email = self.app_context.user_email
        
        # Update status card
        status_layout = self.login_status_card.layout()
        if status_layout and status_layout.count() >= 2:
            status_label = status_layout.itemAt(1).widget()
            if isinstance(status_label, QLabel):
                if is_logged_in:
                    status_text = f"Logged in\n{user_email or 'User'}"
                    status_label.setStyleSheet(f"color: {ThemeColors.status_success()}; font-size: 14px;")
                else:
                    status_text = "Not logged in"
                    status_label.setStyleSheet(f"color: {ThemeColors.status_error()}; font-size: 14px;")
                status_label.setText(status_text)
        
        # Update login button
        if is_logged_in:
            self.login_btn.setText("Re-login")
            self.login_btn.setIcon(FluentIcon.SYNC)
        else:
            self.login_btn.setText("Sign In")
            self.login_btn.setIcon(FluentIcon.PEOPLE)
    
    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        """
        Handle login state changed signal.
        
        Args:
            is_logged_in: Whether user is logged in
        """
        self._update_login_status()
    
    def _on_user_info_changed(self, user_info: dict) -> None:
        """
        Handle user info changed signal.
        
        Args:
            user_info: User information dictionary
        """
        self._update_login_status()
    
    def _on_ap_clicked(self) -> None:
        """Handle AP button click - navigate to AP page."""
        # This will be connected in MainWindow to navigate to AP page
        if hasattr(self, '_navigate_to_ap'):
            self._navigate_to_ap()
    
    def _on_ar_clicked(self) -> None:
        """Handle AR button click - navigate to AR page."""
        # This will be connected in MainWindow to navigate to AR page
        if hasattr(self, '_navigate_to_ar'):
            self._navigate_to_ar()
    
    def _on_login_clicked(self) -> None:
        """Handle login button click - trigger login."""
        # This will be connected in MainWindow to show login dialog
        if hasattr(self, '_trigger_login'):
            self._trigger_login()
    
    def _get_theme_button_text(self) -> str:
        """Get theme button text based on current theme."""
        return "Dark" if not isDarkTheme() else "Light"
    
    def _get_theme_tooltip(self) -> str:
        """Get theme button tooltip based on current theme."""
        return "Switch to light theme" if isDarkTheme() else "Switch to dark theme"
    
    def _on_theme_toggle_clicked(self) -> None:
        """Handle theme toggle button click."""
        # Toggle between dark and light theme
        if isDarkTheme():
            setTheme(Theme.LIGHT)
        else:
            setTheme(Theme.DARK)
        
        # Update button text and tooltip
        self.theme_btn.setText(self._get_theme_button_text())
        self.theme_btn.setToolTip(self._get_theme_tooltip())
        
        # Trigger theme update for all pages after a short delay
        # This ensures the theme change has propagated
        QTimer.singleShot(100, self._trigger_theme_update)
    
    def _trigger_theme_update(self) -> None:
        """Trigger theme update for parent MainWindow."""
        # Find MainWindow through parent hierarchy
        parent = self.parent()
        while parent:
            if hasattr(parent, 'update_theme_for_all_pages'):
                # Use the public method instead
                parent.update_theme_for_all_pages()
                break
            elif hasattr(parent, '_update_all_pages_theme'):
                # Fallback to old method if new one doesn't exist
                try:
                    parent._update_all_pages_theme()
                except Exception:
                    pass
                break
            parent = parent.parent()
    
    def _check_for_updates_background(self) -> None:
        """Check for updates in background without blocking UI."""
        if self.check_update_worker and self.check_update_worker.isRunning():
            return
        
        from workers.update_worker import UpdateCheckWorker
        
        self.check_update_worker = UpdateCheckWorker()
        self.check_update_worker.finished.connect(self._on_update_check_finished)
        self.check_update_worker.start()
    
    def _on_update_check_finished(self, result: dict) -> None:
        """Handle background update check finished."""
        success = result.get("success", False)
        update_available = result.get("update_available", False)
        
        if success and update_available:
            latest_version = result.get("latest_version", "Unknown")
            local_version = result.get("local_version", "Unknown")
            release_info = result.get("release_info", {})
            
            # Show notification dialog to user using qfluentwidgets MessageBox
            msg_box = MessageBox(
                title="Update Available",
                content=f"A new version is available!\n\n"
                f"Current version: {local_version}\n"
                f"Latest version: {latest_version}\n\n"
                f"Would you like to update now?",
                parent=self
            )
            
            # Connect yesSignal to handle user confirmation
            def _handle_update_yes():
                # User chose to update - open update dialog and start download immediately
                from ui.widgets.update_dialog import UpdateDialog
                dialog = UpdateDialog(self)
                # Pre-populate with update info and start download automatically
                dialog._start_update_immediately(release_info, latest_version, local_version)
                dialog.exec()
            
            msg_box.yesSignal.connect(_handle_update_yes)
            msg_box.exec()
    
    def _on_check_update_clicked(self) -> None:
        """Handle check for updates button click - show update dialog."""
        from ui.widgets.update_dialog import UpdateDialog
        dialog = UpdateDialog(self)
        dialog.exec()
    