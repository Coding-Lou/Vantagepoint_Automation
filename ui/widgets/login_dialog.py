"""
Login dialog with modern Fluent Design.

Provides SSO login interface with progress indication.
"""
from PySide6.QtWidgets import QVBoxLayout, QHBoxLayout, QDialog
from PySide6.QtCore import Qt, Signal

from qfluentwidgets import (
    PrimaryPushButton,
    BodyLabel,
    TitleLabel,
    IndeterminateProgressRing,
    FluentIcon,
    CardWidget,
)

from ui.utils.theme_colors import ThemeColors


class LoginDialog(QDialog):
    """
    Modern login dialog for SSO authentication.
    
    Features:
    - Clean Fluent Design style
    - Progress indication during login
    - Error message display
    """
    
    # Signal emitted when login is requested
    login_requested = Signal()
    
    def __init__(self, parent=None):
        """Initialize login dialog."""
        super().__init__(parent)
        self.setWindowTitle("Sign In")
        self.setFixedSize(400, 300)
        
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        """Setup UI layout."""
        # Title
        title_label = TitleLabel("Sign In to QCA Systems")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Description
        desc_label = BodyLabel(
            "Click the button below to sign in using SSO authentication. "
            "A browser window will open for you to complete the login process."
        )
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setStyleSheet(f"color: {ThemeColors.text_muted()};")
        
        # Progress indicator (hidden initially)
        self.progress_ring = IndeterminateProgressRing(self)
        self.progress_ring.setVisible(False)
        
        # Status label
        self.status_label = BodyLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setWordWrap(True)
        self.status_label.setVisible(False)
        
        # Login button
        self.login_btn = PrimaryPushButton("Sign In", self, FluentIcon.PEOPLE)
        self.login_btn.setMinimumHeight(40)
        self.login_btn.clicked.connect(self._on_login_clicked)
        
        # Layout
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)
        
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addStretch()
        layout.addWidget(self.progress_ring, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        layout.addStretch()
        layout.addWidget(self.login_btn)
    
    def _on_login_clicked(self) -> None:
        """Handle login button click."""
        self.login_requested.emit()
        self._set_loading_state(True, "Opening browser for authentication...")
    
    def _set_loading_state(self, loading: bool, message: str = "") -> None:
        """
        Set loading state.
        
        Args:
            loading: Whether login is in progress
            message: Status message to display
        """
        self.login_btn.setEnabled(not loading)
        self.progress_ring.setVisible(loading)
        self.status_label.setVisible(bool(message))
        
        if message:
            self.status_label.setText(message)
            if "error" in message.lower() or "failed" in message.lower():
                self.status_label.setStyleSheet(f"color: {ThemeColors.status_error()};")
            elif "success" in message.lower():
                self.status_label.setStyleSheet(f"color: {ThemeColors.status_success()};")
            else:
                self.status_label.setStyleSheet(f"color: {ThemeColors.text_muted()};")
    
    def show_success(self) -> None:
        """Show success message and close dialog."""
        self._set_loading_state(False, "Login successful!")
        # Close dialog after a short delay
        from PySide6.QtCore import QTimer
        QTimer.singleShot(1000, self.accept)
    
    def show_error(self, error_message: str) -> None:
        """
        Show error message.
        
        Args:
            error_message: Error message to display
        """
        self._set_loading_state(False, f"Login failed: {error_message}")
        self.login_btn.setEnabled(True)
