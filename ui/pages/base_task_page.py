"""
Base task page class with modern Fluent Design.

Provides common structure for all task pages:
- Top: Page title and description
- Middle: Configuration area (Card-based)
- Bottom: Execution log area (Card-based)
- Action button: Primary button in bottom-right
"""
from typing import Dict, Any, Optional, Tuple
from datetime import datetime
from pathlib import Path
import json
import time

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTextEdit,
    QScrollArea,
)
from PySide6.QtCore import Signal, Slot, Qt

from qfluentwidgets import (
    PrimaryPushButton,
    CardWidget,
    TitleLabel,
    BodyLabel,
    PlainTextEdit,
    PushButton,
)

from workers.base_worker import BaseWorker

#region agent log
DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Optional[Dict[str, Any]] = None,
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
#endregion


class BaseTaskPage(QWidget):
    """
    Base class for all task pages.
    
    Provides common structure:
    - Top: Configuration area (dynamically generated)
    - Middle: Action buttons
    - Bottom: Execution log area
    """
    
    # Signals
    task_started = Signal(str)  # task_id
    task_finished = Signal(str, dict)  # task_id, result
    task_failed = Signal(str, str)  # task_id, error_message
    
    # Log level colors
    LOG_COLOR_INFO = "#d4d4d4"
    LOG_COLOR_SUCCESS = "#4ec9b0"
    LOG_COLOR_WARNING = "#dcdcaa"
    LOG_COLOR_ERROR = "#f48771"
    
    def __init__(self, task_id: str, task_name: str, parent=None):
        """
        Initialize task page.
        
        Args:
            task_id: Task identifier
            task_name: Task display name
            parent: Parent widget
        """
        super().__init__(parent)
        self.task_id = task_id
        self.task_name = task_name
        self.current_worker: Optional[BaseWorker] = None
        #region agent log
        _agent_log(
            "H4",
            "BaseTaskPage.__init__",
            "Task page initializing",
            {"task_id": task_id, "task_name": task_name},
        )
        #endregion
        self._setup_ui()
        self._load_config()
    
    def _setup_ui(self) -> None:
        """Setup UI layout with modern Fluent Design."""
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        # === Scrollable content (so Execute button is always visible) ===
        self._scroll = QScrollArea(self)
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(16)

        # 1. Page header (title and description)
        self._setup_page_header(content_layout)

        # 2. Configuration area (Card-based)
        self.config_card = CardWidget(content)
        self.config_layout = QVBoxLayout()
        self.config_layout.setContentsMargins(16, 16, 16, 16)
        self.config_layout.setSpacing(12)
        self.config_card.setLayout(self.config_layout)
        content_layout.addWidget(self.config_card)

        # Create configuration widgets (implemented by subclass)
        self._create_config_widgets()

        # 3. Log area (Card-based)
        self.log_card = CardWidget(content)
        log_card_layout = QVBoxLayout()
        log_card_layout.setContentsMargins(16, 16, 16, 16)
        log_card_layout.setSpacing(12)

        # Log title
        log_title = BodyLabel("Execution Log")
        log_title.setStyleSheet("font-weight: 600;")
        log_card_layout.addWidget(log_title)

        # Log viewer - Use QTextEdit for HTML support
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        # Set minimum height for better log visibility
        self.log_viewer.setMinimumHeight(500)
        self.log_viewer.setStyleSheet(f"""
            QTextEdit {{
                background-color: #1e1e1e;
                color: {self.LOG_COLOR_INFO};
                font-family: 'Consolas', 'Courier New', monospace;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
                padding: 8px;
            }}
        """)
        log_card_layout.addWidget(self.log_viewer)

        # Log actions
        log_actions_layout = QHBoxLayout()
        log_actions_layout.addStretch()

        clear_btn = PushButton("Clear Log", content)
        clear_btn.clicked.connect(self.log_viewer.clear)
        log_actions_layout.addWidget(clear_btn)

        log_card_layout.addLayout(log_actions_layout)
        self.log_card.setLayout(log_card_layout)
        content_layout.addWidget(self.log_card)

        # Spacer so the bottom of scroll content isn't glued to the action bar
        content_layout.addStretch(1)

        self._scroll.setWidget(content)
        root.addWidget(self._scroll, stretch=1)

        # === Fixed action bar (always visible) ===
        action_layout = QHBoxLayout()
        action_layout.addStretch()

        self.cancel_btn = PushButton("Cancel", self)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._on_cancel_clicked)
        action_layout.addWidget(self.cancel_btn)

        self.execute_btn = PrimaryPushButton("Execute", self)
        self.execute_btn.setMinimumWidth(120)
        self.execute_btn.setMinimumHeight(36)
        self.execute_btn.clicked.connect(self._on_execute_clicked)
        action_layout.addWidget(self.execute_btn)

        root.addLayout(action_layout)
    
    def _setup_page_header(self, parent_layout: QVBoxLayout) -> None:
        """
        Setup page header with title and description.
        
        Args:
            parent_layout: Parent layout to add header to
        """
        # Title
        title_label = TitleLabel(self.task_name)
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        parent_layout.addWidget(title_label)
        
        # Description (can be overridden by subclasses)
        description = self._get_page_description()
        if description:
            desc_label = BodyLabel(description)
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #808080;")
            parent_layout.addWidget(desc_label)
    
    def _get_page_description(self) -> str:
        """
        Get page description text.
        
        Returns:
            Description text (empty string by default)
        """
        return ""
    
    def _create_config_widgets(self) -> None:
        """
        Create configuration widgets.
        
        This method must be implemented by subclasses.
        """
        raise NotImplementedError("Subclass must implement _create_config_widgets()")
    
    def _load_config(self) -> None:
        """
        Load configuration from config file.
        
        This method can be overridden by subclasses to load default values.
        """
        pass
    
    def _validate_params(self) -> Tuple[bool, str]:
        """
        Validate task parameters.
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        return True, ""
    
    def _get_params(self) -> Dict[str, Any]:
        """
        Get task parameters from UI.
        
        Returns:
            Parameters dictionary
        """
        return {}
    
    def _create_worker(self, params: Dict[str, Any]) -> BaseWorker:
        """
        Create task worker.
        
        Args:
            params: Task parameters
            
        Returns:
            Worker instance
        """
        raise NotImplementedError("Subclass must implement _create_worker()")
    
    @Slot()
    def _on_execute_clicked(self) -> None:
        """Handle execute button click."""
        # Validate parameters
        is_valid, error_msg = self._validate_params()
        if not is_valid:
            self._append_log("ERROR", f"Parameter validation failed: {error_msg}")
            return
        
        #region agent log
        _agent_log(
            "H1",
            "BaseTaskPage._on_execute_clicked",
            "Execute clicked - checking existing worker",
            {"has_worker": bool(self.current_worker), "is_running": self.current_worker.isRunning() if self.current_worker else False},
        )
        #endregion
        
        # Clean up any existing worker before creating a new one
        if self.current_worker:
            #region agent log
            _agent_log(
                "H1",
                "BaseTaskPage._on_execute_clicked",
                "Existing worker found",
                {"is_running": self.current_worker.isRunning()},
            )
            #endregion
            if self.current_worker.isRunning():
                self._append_log("WARNING", "Task is already running. Please wait for completion.")
                #region agent log
                _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Worker still running - blocking execution", {})
                #endregion
                return
            else:
                # Worker exists but not running - clean it up
                worker = self.current_worker
                try:
                    worker.log_signal.disconnect()
                except Exception:
                    pass
                try:
                    worker.progress_signal.disconnect()
                except Exception:
                    pass
                try:
                    worker.finished.disconnect()
                except Exception:
                    pass
                self.current_worker = None
                #region agent log
                _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Old worker cleaned up", {})
                #endregion
        
        # Get parameters
        params = self._get_params()
        #region agent log
        _agent_log(
            "H3",
            "BaseTaskPage._on_execute_clicked",
            "Parameters collected for execution",
            {"task_id": self.task_id, "params": params},
        )
        #endregion
        
        # Create worker
        #region agent log
        _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Creating new worker", {"task_id": self.task_id})
        #endregion
        self.current_worker = self._create_worker(params)
        
        # Connect signals
        self.current_worker.log_signal.connect(self._on_log_received)
        self.current_worker.progress_signal.connect(self._on_progress_received)
        self.current_worker.finished.connect(self._on_worker_finished)
        
        #region agent log
        _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Worker created and signals connected", {"task_id": self.task_id})
        #endregion
        
        # Update UI state
        self._set_execution_enabled(False)
        self.log_viewer.clear()
        self._append_log("INFO", f"Starting task: {self.task_name}")
        
        # Emit task started signal
        self.task_started.emit(self.task_id)
        
        # Start worker (non-blocking)
        #region agent log
        _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Starting worker thread", {"task_id": self.task_id})
        #endregion
        self.current_worker.start()
        #region agent log
        _agent_log("H1", "BaseTaskPage._on_execute_clicked", "Worker thread started", {"task_id": self.task_id, "is_running": self.current_worker.isRunning()})
        #endregion
    
    @Slot()
    def _on_cancel_clicked(self) -> None:
        """Handle cancel button click."""
        #region agent log
        _agent_log(
            "H1",
            "BaseTaskPage._on_cancel_clicked",
            "Cancel clicked",
            {"has_worker": bool(self.current_worker), "is_running": self.current_worker.isRunning() if self.current_worker else False},
        )
        #endregion
        if self.current_worker:
            worker = self.current_worker  # Store reference before clearing
            was_running = worker.isRunning()
            
            if was_running:
                worker.cancel()
                self._append_log("INFO", "Cancelling task...")
                # Disconnect signals BEFORE termination to prevent stale callbacks
                try:
                    worker.log_signal.disconnect()
                except Exception:
                    pass
                try:
                    worker.progress_signal.disconnect()
                except Exception:
                    pass
                try:
                    worker.finished.disconnect()
                except Exception:
                    pass
                
                # Request thread termination (forceful if needed)
                worker.terminate()
                # Wait for thread to finish (with timeout)
                if not worker.wait(2000):  # Wait up to 2 seconds
                    # If still running, force terminate again
                    worker.terminate()
                    worker.wait(1000)  # Final wait
            
            # Clear worker reference immediately after disconnecting
            self.current_worker = None
            #region agent log
            _agent_log("H1", "BaseTaskPage._on_cancel_clicked", "Worker cleaned up", {"was_running": was_running})
            #endregion
            
            # Update UI state immediately
            self._set_execution_enabled(True)
            self._append_log("INFO", "Task cancelled")
            
            # Emit task_failed signal to notify MainWindow to re-enable all pages
            # This ensures that other pages (AR, etc.) are also re-enabled after cancellation
            if was_running:
                self.task_failed.emit(self.task_id, "Task was cancelled by user")
    
    @Slot(dict)
    def _on_worker_finished(self, result: Dict[str, Any]) -> None:
        """
        Handle worker finished signal.
        
        Args:
            result: Result dictionary
        """
        #region agent log
        _agent_log(
            "H1",
            "BaseTaskPage._on_worker_finished",
            "Worker finished signal received",
            {"has_worker": bool(self.current_worker), "success": bool(result.get("success"))},
        )
        #endregion
        # Only process if this is the current worker (ignore stale signals from cancelled workers)
        if not self.current_worker:
            #region agent log
            _agent_log("H1", "BaseTaskPage._on_worker_finished", "Ignoring stale signal - no current worker", {})
            #endregion
            return
        
        # Store worker reference before clearing
        worker = self.current_worker
        self.current_worker = None
        
        # Disconnect signals to prevent any further callbacks
        try:
            worker.log_signal.disconnect()
        except Exception:
            pass
        try:
            worker.progress_signal.disconnect()
        except Exception:
            pass
        try:
            worker.finished.disconnect()
        except Exception:
            pass
        
        # Update UI state
        self._set_execution_enabled(True)
        
        if result.get("success"):
            self._append_log("SUCCESS", "Task completed successfully")
            self.task_finished.emit(self.task_id, result)
        else:
            error_msg = result.get("message", "Unknown error")
            # Don't log error for cancelled tasks
            if "cancelled" not in error_msg.lower():
                self._append_log("ERROR", f"Task failed: {error_msg}")
            self.task_failed.emit(self.task_id, error_msg)
        
        #region agent log
        _agent_log("H1", "BaseTaskPage._on_worker_finished", "Worker finished processed and cleaned up", {})
        #endregion
    
    @Slot(str, str)
    def _on_log_received(self, level: str, message: str) -> None:
        """
        Handle log signal from worker.
        
        Args:
            level: Log level
            message: Log message
        """
        self._append_log(level, message)
    
    @Slot(int, int, str)
    def _on_progress_received(self, current: int, total: int, message: str) -> None:
        """
        Handle progress signal from worker.
        
        Args:
            current: Current progress value
            total: Total progress value
            message: Progress message
        """
        if total > 0:
            progress_text = f"Progress: {current}/{total} ({current*100//total}%)"
            if message:
                progress_text += f" - {message}"
            self._append_log("INFO", progress_text)
    
    def _append_log(self, level: str, message: str) -> None:
        """
        Append log message to log viewer.
        
        Args:
            level: Log level (INFO, SUCCESS, WARNING, ERROR)
            message: Log message
        """
        # Format timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Get color based on level
        color_map = {
            "INFO": self.LOG_COLOR_INFO,
            "SUCCESS": self.LOG_COLOR_SUCCESS,
            "WARNING": self.LOG_COLOR_WARNING,
            "ERROR": self.LOG_COLOR_ERROR,
        }
        color = color_map.get(level, self.LOG_COLOR_INFO)
        
        # Format message
        formatted_msg = f'<span style="color: {color}">[{timestamp}] [{level}] {message}</span><br>'
        
        # Append to log viewer
        self.log_viewer.append(formatted_msg)
        
        # Auto scroll to bottom
        scrollbar = self.log_viewer.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def _set_execution_enabled(self, enabled: bool) -> None:
        """
        Set execution button enabled state.
        
        Args:
            enabled: Whether execution is enabled
        """
        self.execute_btn.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)
