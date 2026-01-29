"""
Update dialog for checking and performing updates.
"""
from typing import Optional
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTextEdit,
)
from PySide6.QtCore import Qt

from qfluentwidgets import (
    PrimaryPushButton,
    PushButton,
    BodyLabel,
    TitleLabel,
    CardWidget,
    FluentIcon,
)

from workers.update_worker import UpdateWorker, UpdateCheckWorker


class UpdateDialog(QDialog):
    """
    Dialog for checking and performing updates.
    
    Features:
    - Check for updates
    - Display update information
    - Download and install updates with progress
    """
    
    def __init__(self, parent=None):
        """Initialize update dialog."""
        super().__init__(parent)
        self.setWindowTitle("Check for Updates")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        self.current_worker: Optional[UpdateWorker] = None
        self.check_worker: Optional[UpdateCheckWorker] = None
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup dialog UI."""
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = TitleLabel("Update")
        layout.addWidget(title)
        
        # Status card
        self.status_card = CardWidget(self)
        status_layout = QVBoxLayout(self.status_card)
        status_layout.setContentsMargins(16, 16, 16, 16)
        status_layout.setSpacing(8)
        
        self.status_label = BodyLabel("Click 'Check for Updates' to check for new versions.")
        self.status_label.setWordWrap(True)
        status_layout.addWidget(self.status_label)
        
        self.version_label = BodyLabel("")
        self.version_label.setStyleSheet("color: #808080;")
        status_layout.addWidget(self.version_label)
        
        layout.addWidget(self.status_card)
        
        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        layout.addWidget(self.progress_bar)
        
        # Log area
        log_label = BodyLabel("Update Log")
        log_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(log_label)
        
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMinimumHeight(150)
        self.log_viewer.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: 'Consolas', 'Courier New', monospace;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        layout.addWidget(self.log_viewer)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.check_btn = PrimaryPushButton("Check for Updates", self, FluentIcon.SYNC)
        self.check_btn.clicked.connect(self._on_check_clicked)
        button_layout.addWidget(self.check_btn)
        
        self.update_btn = PrimaryPushButton("Download and Install", self, FluentIcon.DOWNLOAD)
        self.update_btn.setVisible(False)
        self.update_btn.clicked.connect(self._on_update_clicked)
        button_layout.addWidget(self.update_btn)
        
        self.cancel_btn = PushButton("Cancel", self)
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        # Store update info
        self.update_info: Optional[dict] = None
        self.auto_start_download = False
    
    def _on_check_clicked(self):
        """Handle check for updates button click."""
        if self.check_worker and self.check_worker.isRunning():
            return
        
        self.check_btn.setEnabled(False)
        self.status_label.setText("Checking for updates...")
        self.version_label.setText("")
        self.log_viewer.clear()
        self._append_log("INFO", "Checking for updates...")
        
        # Create and start check worker
        self.check_worker = UpdateCheckWorker()
        self.check_worker.log_signal.connect(self._on_log_received)
        self.check_worker.finished.connect(self._on_check_finished)
        self.check_worker.start()
    
    def _start_update_immediately(self, release_info: dict, latest_version: str, local_version: str):
        """
        Start update download immediately without user clicking button.
        
        Args:
            release_info: Release information dictionary
            latest_version: Latest version string
            local_version: Local version string
        """
        self.update_info = {
            "latest_version": latest_version,
            "local_version": local_version,
            "release_info": release_info,
        }
        self.status_label.setText(f"Update available: {latest_version}")
        self.version_label.setText(f"Current version: {local_version} → Latest version: {latest_version}")
        self.update_btn.setVisible(True)
        self.check_btn.setEnabled(False)
        # Automatically start download
        self._start_download()
    
    def _on_update_clicked(self):
        """Handle download and install button click."""
        if not self.update_info:
            return
        self._start_download()
    
    def _start_download(self):
        """Start the download process."""
        if self.current_worker and self.current_worker.isRunning():
            return
        
        self.check_btn.setEnabled(False)
        self.update_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_viewer.clear()
        self._append_log("INFO", "Starting update download...")
        self.status_label.setText("Downloading update...")
        
        # Create and start update worker
        self.current_worker = UpdateWorker()
        self.current_worker.log_signal.connect(self._on_log_received)
        self.current_worker.progress_signal.connect(self._on_progress_received)
        self.current_worker.finished.connect(self._on_update_finished)
        self.current_worker.start()
    
    def _on_check_finished(self, result: dict):
        """Handle update check finished."""
        self.check_btn.setEnabled(True)
        
        success = result.get("success", False)
        update_available = result.get("update_available", False)
        
        if success and update_available:
            latest_version = result.get("latest_version", "Unknown")
            local_version = result.get("local_version", "Unknown")
            
            self.status_label.setText(f"Update available: {latest_version}")
            self.version_label.setText(f"Current version: {local_version} → Latest version: {latest_version}")
            self.update_btn.setVisible(True)
            self.update_info = result
            self._append_log("SUCCESS", f"Update available: {latest_version}")
        elif success:
            local_version = result.get("local_version", "Unknown")
            self.status_label.setText("You are using the latest version")
            self.version_label.setText(f"Current version: {local_version}")
            self.update_btn.setVisible(False)
            self._append_log("SUCCESS", "Already on latest version")
        else:
            error_msg = result.get("message", "Unknown error")
            self.status_label.setText(f"Update check failed: {error_msg}")
            self.update_btn.setVisible(False)
            self._append_log("ERROR", f"Update check failed: {error_msg}")
    
    def _on_update_finished(self, result: dict):
        """Handle update finished."""
        success = result.get("success", False)
        cancelled = result.get("cancelled", False)
        
        if cancelled:
            self.check_btn.setEnabled(True)
            self.update_btn.setEnabled(True)
            self.progress_bar.setVisible(False)
            self._append_log("INFO", "Update cancelled")
        elif success:
            # Update will restart the application, so we can close the dialog
            self._append_log("SUCCESS", "Update completed. Application will restart.")
            # Close dialog and exit application to allow updater to replace exe
            self.accept()
            # Close the main application
            from PySide6.QtWidgets import QApplication
            app = QApplication.instance()
            if app:
                # Give a small delay to ensure log message is displayed
                from PySide6.QtCore import QTimer
                QTimer.singleShot(500, app.quit)
        else:
            error_msg = result.get("message", "Unknown error")
            self.check_btn.setEnabled(True)
            self.update_btn.setEnabled(True)
            self.progress_bar.setVisible(False)
            self._append_log("ERROR", f"Update failed: {error_msg}")
    
    def _on_log_received(self, level: str, message: str):
        """Handle log signal from worker."""
        self._append_log(level, message)
    
    def _on_progress_received(self, current: int, total: int):
        """Handle progress signal from worker - update progress bar only, no log."""
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
            # Update status label with progress percentage
            percent = int((current / total) * 100)
            self.status_label.setText(f"Downloading update... {percent}% ({current / 1024 / 1024:.2f} MB / {total / 1024 / 1024:.2f} MB)")
    
    def _append_log(self, level: str, message: str):
        """Append log message to log viewer."""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        color_map = {
            "INFO": "#d4d4d4",
            "SUCCESS": "#4ec9b0",
            "WARNING": "#dcdcaa",
            "ERROR": "#f48771",
        }
        color = color_map.get(level, "#d4d4d4")
        
        formatted_msg = f'<span style="color: {color}">[{timestamp}] [{level}] {message}</span><br>'
        self.log_viewer.append(formatted_msg)
        
        # Auto scroll to bottom
        scrollbar = self.log_viewer.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
