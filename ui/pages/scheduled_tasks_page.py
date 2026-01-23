"""
Scheduled Tasks page for managing scheduled task configurations.

Features:
- Display list of scheduled tasks from config/tasks.json
- Edit task configurations
- Manual task execution
- Initialize scheduled tasks in Windows Task Scheduler
"""
from typing import Dict, Any, List, Optional
from pathlib import Path
import json
from datetime import datetime
from ui.utils.log_redirector import LogRedirector

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QPushButton,
    QTextEdit,
    QScrollArea,
    QLineEdit,
    QFormLayout,
    QDialog,
    QDialogButtonBox,
    QMessageBox,
    QComboBox,
    QTimeEdit,
    QCheckBox,
    QGroupBox,
)
from PySide6.QtCore import Qt, Signal, Slot, QThread, QTime

from qfluentwidgets import (
    CardWidget,
    TitleLabel,
    BodyLabel,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
    PlainTextEdit,
    LineEdit,
)

import tools.schedule_task as schedule_task_module
import tools.util as util_module
import tools.login as login_module
import tools.config_manager as config_manager

# Core modules are imported lazily in TaskExecutionWorker to avoid import errors


def get_tasks_file() -> Path:
    """
    Get path to legacy tasks.json file used by schedule_task module.
    
    This is kept for backward compatibility so that we can still
    generate the tasks.json file expected by tools.schedule_task.
    """
    base_dir = Path(__file__).resolve().parent.parent.parent
    # Prefer the same resolution logic as schedule_task.TASKS_FILE
    try:
        return schedule_task_module.TASKS_FILE  # type: ignore[attr-defined]
    except Exception:
        # Fallback: config/tasks.json in development
        return base_dir / "config" / "tasks.json"


def load_tasks() -> List[Dict[str, Any]]:
    """
    Load scheduled tasks from config.json (schedule_tasks key).
    
    Falls back to legacy task.json/tasks.json if needed for backward compatibility.
    """
    # Primary source: config.json -> schedule_tasks
    try:
        config_tasks = config_manager.get_config_value(["schedule_tasks"])
        if isinstance(config_tasks, list):
            return config_tasks
        if isinstance(config_tasks, dict):
            # If stored as a single dict, wrap in list
            return [config_tasks]
    except Exception as e:
        print(f"Error loading schedule_tasks from config.json: {e}")
    
    # Fallback: legacy tasks.json structure
    tasks_file = get_tasks_file()
    if not tasks_file.exists():
        return []
    
    try:
        with open(tasks_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("tasks", [])
    except Exception as e:
        print(f"Error loading tasks from legacy file: {e}")
        return []


def save_tasks(tasks: List[Dict[str, Any]]) -> bool:
    """
    Save scheduled tasks to config.json under \"schedule_tasks\" key.
    
    This function performs a merge/update operation:
    - Reads the entire config.json file
    - Updates only the schedule_tasks field
    - Preserves all other existing fields (AP, AR, etc.)
    - Maintains the original key name \"schedule_tasks\"
    - Saves the complete configuration back to file
    
    Also writes a legacy tasks.json file so that tools.schedule_task continues to work.
    
    Args:
        tasks: List of task dictionaries to save
        
    Returns:
        True if save was successful, False otherwise
    """
    config_path = config_manager.get_config_path()
    success = False
    
    try:
        # Read entire config.json to preserve all other values
        with open(config_path, "r", encoding="utf-8") as f:
            original_content = f.read()
            config = json.loads(original_content)
        
        # Detect original indentation (2 or 4 spaces)
        # Check first non-empty line after opening brace
        indent_size = 4  # Default
        for line in original_content.split('\n'):
            stripped = line.lstrip()
            if stripped and stripped.startswith('"'):
                # Calculate indent
                indent = len(line) - len(stripped)
                if indent > 0:
                    indent_size = indent
                    break
        
        # Update only the schedule_tasks field (merge operation)
        # This preserves all other fields like AP, AR, etc.
        config["schedule_tasks"] = tasks
        
        # Save back with detected indentation to maintain original format
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=indent_size, ensure_ascii=False)
            f.flush()
        
        success = True
    except FileNotFoundError:
        print(f"Error: config.json not found at {config_path}")
        success = False
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in config.json: {e}")
        success = False
    except Exception as e:
        print(f"Error saving schedule_tasks to config.json: {e}")
        success = False
    
    return success


def format_schedule_info(schedule: Dict[str, Any]) -> str:
    """Format schedule information for display."""
    stype = schedule.get("type", "")
    if stype == "minute":
        interval = schedule.get("interval", 5)
        return f"Every {interval} minute(s)"
    elif stype == "daily":
        time_str = schedule.get("time", "")
        return f"Daily at {time_str}"
    elif stype == "weekly":
        days = schedule.get("days", [])
        time_str = schedule.get("time", "")
        days_str = ", ".join(days)
        return f"Weekly on {days_str} at {time_str}"
    else:
        return str(schedule)


def parse_task_args(args: str) -> Dict[str, Any]:
    """Parse task args string to determine execution parameters."""
    parsed = {}
    if not args:
        return parsed
    
    # Parse arguments
    parts = args.split()
    i = 0
    while i < len(parts):
        arg = parts[i]
        if arg.startswith("--"):
            # Handle --key=value format
            if "=" in arg:
                key, value = arg[2:].split("=", 1)
                parsed[key] = value
                i += 1
            else:
                key = arg[2:]
                if i + 1 < len(parts) and not parts[i + 1].startswith("--"):
                    parsed[key] = parts[i + 1]
                    i += 2
                else:
                    parsed[key] = True
                    i += 1
        else:
            i += 1
    
    return parsed


class TaskExecutionWorker(QThread):
    """
    Worker thread for executing scheduled tasks manually.
    
    Captures stdout/stderr output from task execution and redirects to UI log viewer.
    """
    
    log_signal = Signal(str, str)  # level, message
    finished = Signal(dict)  # result
    
    def __init__(self, task: Dict[str, Any]):
        """Initialize worker with task configuration."""
        super().__init__()
        self.task = task
        self.log_redirector: Optional[LogRedirector] = None
        self.old_stdout: Optional[Any] = None
        self.old_stderr: Optional[Any] = None
    
    def _setup_log_redirect(self) -> None:
        """Setup stdout/stderr redirection to Signal."""
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        self.log_redirector = LogRedirector(self.log_signal.emit)
        sys.stdout = self.log_redirector
        sys.stderr = self.log_redirector
    
    def _restore_log_redirect(self) -> None:
        """Restore original stdout/stderr."""
        if self.old_stdout:
            sys.stdout = self.old_stdout
        if self.old_stderr:
            sys.stderr = self.old_stderr
    
    def run(self):
        """Execute task in background thread with log redirection."""
        try:
            # Setup log redirection to capture print() output
            self._setup_log_redirect()
            
            args_str = self.task.get("args", "")
            parsed_args = parse_task_args(args_str)
            
            # Check login status
            login = util_module.check_login()
            if not login:
                self.log_signal.emit("INFO", "Not logged in, attempting login...")
                while not login:
                    login_module.sso_login()
                    login = util_module.check_login()
            
            # Execute based on parsed args (lazy import to avoid import errors)
            if parsed_args.get("ap"):
                self.log_signal.emit("INFO", "Executing AP task...")
                import core.ap as ap_module
                ap_module.main()
            elif parsed_args.get("ar"):
                self.log_signal.emit("INFO", "Executing AR task...")
                import core.ar as ar_module
                ar_module.main()
            elif parsed_args.get("project_status"):
                self.log_signal.emit("INFO", "Executing Project Status task...")
                import core.project_status as project_status_module
                project_status_module.main()
            elif parsed_args.get("bridge"):
                self.log_signal.emit("INFO", "Executing Bridge Report task...")
                import core.bridge_report as bridge_report_module
                bridge_report_module.main()
            elif parsed_args.get("shipping"):
                self.log_signal.emit("INFO", "Executing Shipping Monitor task...")
                import core.shipping_monitor as shipping_monitor_module
                shipping_monitor_module.main()
            elif parsed_args.get("daily_received"):
                self.log_signal.emit("INFO", "Executing Daily Received task...")
                import core.daily_received as daily_received_module
                daily_received_module.main()
            elif parsed_args.get("on_call"):
                vendor_name = parsed_args.get("name")
                if vendor_name:
                    self.log_signal.emit("INFO", f"Executing On Call task for {vendor_name}...")
                    import core.on_call as on_call_module
                    on_call_module.run_oncall_task(vendorName=vendor_name)
                else:
                    self.log_signal.emit("ERROR", "On Call task requires --name parameter")
                    self.finished.emit({"success": False, "message": "Missing --name parameter"})
                    return
            else:
                self.log_signal.emit("ERROR", f"Unknown task type: {args_str}")
                self.finished.emit({"success": False, "message": f"Unknown task type: {args_str}"})
                return
            
            self.log_signal.emit("SUCCESS", "Task executed successfully")
            self.finished.emit({"success": True, "message": "Task completed"})
        except Exception as e:
            self.log_signal.emit("ERROR", f"Task execution failed: {str(e)}")
            self.finished.emit({"success": False, "message": str(e)})
        finally:
            # Restore log redirection
            self._restore_log_redirect()


class TaskEditDialog(QDialog):
    """Dialog for editing task configuration with Windows Task Scheduler-like interface."""
    
    def __init__(self, task: Dict[str, Any], parent=None):
        """Initialize edit dialog with task data."""
        super().__init__(parent)
        self.task = task.copy()
        self.setWindowTitle("Edit Scheduled Task")
        self.setMinimumWidth(650)
        self.setMinimumHeight(500)
        self._setup_ui()
        self._load_task_data()
    
    def _setup_ui(self):
        """Setup dialog UI with Windows Task Scheduler-like interface."""
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        
        # General settings group
        general_group = QGroupBox("General")
        general_layout = QFormLayout()
        general_layout.setSpacing(12)
        
        # Task name
        self.id_edit = LineEdit()
        general_layout.addRow("Task Name:", self.id_edit)
        
        # Executable (read-only, for reference)
        self.exe_edit = LineEdit()
        self.exe_edit.setReadOnly(True)
        self.exe_edit.setStyleSheet("background-color: #f5f5f5;")
        general_layout.addRow("Executable:", self.exe_edit)
        
        # Arguments
        self.args_edit = LineEdit()
        self.args_edit.setPlaceholderText("e.g., --daily_received or --on_call --name=Pembina")
        general_layout.addRow("Arguments:", self.args_edit)
        
        general_group.setLayout(general_layout)
        layout.addWidget(general_group)
        
        # Trigger settings group
        trigger_group = QGroupBox("Trigger")
        trigger_layout = QVBoxLayout()
        trigger_layout.setSpacing(12)
        
        # Schedule type
        type_layout = QHBoxLayout()
        type_layout.addWidget(BodyLabel("Schedule Type:"))
        self.schedule_type_combo = QComboBox()
        self.schedule_type_combo.addItems(["Daily", "Weekly", "Minute"])
        self.schedule_type_combo.currentTextChanged.connect(self._on_schedule_type_changed)
        type_layout.addWidget(self.schedule_type_combo)
        type_layout.addStretch()
        trigger_layout.addLayout(type_layout)
        
        # Time picker
        time_layout = QHBoxLayout()
        time_layout.addWidget(BodyLabel("Start Time:"))
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime(9, 0))
        time_layout.addWidget(self.time_edit)
        time_layout.addStretch()
        trigger_layout.addLayout(time_layout)
        
        # Schedule-specific options container
        self.schedule_options_widget = QWidget()
        self.schedule_options_layout = QVBoxLayout(self.schedule_options_widget)
        self.schedule_options_layout.setContentsMargins(0, 0, 0, 0)
        trigger_layout.addWidget(self.schedule_options_widget)
        
        # Daily options
        self.daily_options = QWidget()
        daily_layout = QHBoxLayout(self.daily_options)
        daily_layout.addWidget(BodyLabel("Repeat Every:"))
        self.daily_interval_spin = QComboBox()
        self.daily_interval_spin.addItems(["1", "2", "3", "4", "5", "6", "7", "14", "21", "30"])
        daily_layout.addWidget(self.daily_interval_spin)
        daily_layout.addWidget(BodyLabel("day(s)"))
        daily_layout.addStretch()
        
        # Weekly options
        self.weekly_options = QWidget()
        weekly_layout = QVBoxLayout(self.weekly_options)
        weekly_layout.setSpacing(8)
        weekly_layout.addWidget(BodyLabel("Repeat On:"))
        days_layout = QHBoxLayout()
        self.day_checkboxes = {}
        days = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
        day_labels = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for day, label in zip(days, day_labels):
            checkbox = QCheckBox(label)
            self.day_checkboxes[day] = checkbox
            days_layout.addWidget(checkbox)
        days_layout.addStretch()
        weekly_layout.addLayout(days_layout)
        
        # Minute options
        self.minute_options = QWidget()
        minute_layout = QHBoxLayout(self.minute_options)
        minute_layout.addWidget(BodyLabel("Repeat Every:"))
        self.minute_interval_spin = QComboBox()
        self.minute_interval_spin.addItems(["1", "5", "10", "15", "30", "60"])
        minute_layout.addWidget(self.minute_interval_spin)
        minute_layout.addWidget(BodyLabel("minute(s)"))
        minute_layout.addStretch()
        
        # Add all options to container (initially hidden)
        self.schedule_options_layout.addWidget(self.daily_options)
        self.schedule_options_layout.addWidget(self.weekly_options)
        self.schedule_options_layout.addWidget(self.minute_options)
        self.schedule_options_layout.addStretch()
        
        trigger_group.setLayout(trigger_layout)
        layout.addWidget(trigger_group)
        
        layout.addStretch()
        
        # Buttons: Save and Cancel
        buttons = QHBoxLayout()
        buttons.addStretch()
        
        cancel_btn = PushButton("Cancel", self)
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(cancel_btn)
        
        save_btn = PrimaryPushButton("Save", self, FluentIcon.SAVE)
        save_btn.clicked.connect(self._validate_and_save)
        buttons.addWidget(save_btn)
        
        layout.addLayout(buttons)
    
    def _on_schedule_type_changed(self, text: str):
        """Handle schedule type change."""
        # Hide all options
        self.daily_options.setVisible(False)
        self.weekly_options.setVisible(False)
        self.minute_options.setVisible(False)
        
        # Show relevant options
        if text == "Daily":
            self.daily_options.setVisible(True)
        elif text == "Weekly":
            self.weekly_options.setVisible(True)
        elif text == "Minute":
            self.minute_options.setVisible(True)
    
    def _load_task_data(self):
        """Load task data into form."""
        self.id_edit.setText(self.task.get("id", ""))
        self.exe_edit.setText(self.task.get("exe", ""))
        self.args_edit.setText(self.task.get("args", ""))
        
        schedule = self.task.get("schedule", {})
        schedule_type = schedule.get("type", "daily")
        
        # Set schedule type
        if schedule_type == "daily":
            self.schedule_type_combo.setCurrentText("Daily")
        elif schedule_type == "weekly":
            self.schedule_type_combo.setCurrentText("Weekly")
        elif schedule_type == "minute":
            self.schedule_type_combo.setCurrentText("Minute")
        
        # Set time
        time_str = schedule.get("time", "09:00")
        try:
            hour, minute = map(int, time_str.split(":"))
            self.time_edit.setTime(QTime(hour, minute))
        except Exception:
            pass
        
        # Set schedule-specific options
        if schedule_type == "daily":
            interval = str(schedule.get("interval", 1))
            if interval in [str(i) for i in range(1, 32)]:
                self.daily_interval_spin.setCurrentText(interval)
        elif schedule_type == "weekly":
            days = schedule.get("days", [])
            for day, checkbox in self.day_checkboxes.items():
                checkbox.setChecked(day in days)
        elif schedule_type == "minute":
            interval = str(schedule.get("interval", 5))
            if interval in ["1", "5", "10", "15", "30", "60"]:
                self.minute_interval_spin.setCurrentText(interval)
        
        # Trigger UI update
        self._on_schedule_type_changed(self.schedule_type_combo.currentText())
    
    def _validate_and_save(self):
        """Validate form data before saving and closing dialog."""
        if not self.id_edit.text().strip():
            QMessageBox.warning(self, "Validation Error", "Task Name cannot be empty.")
            return
        
        if self.schedule_type_combo.currentText() == "Weekly":
            selected_days = [day for day, checkbox in self.day_checkboxes.items() if checkbox.isChecked()]
            if not selected_days:
                QMessageBox.warning(self, "Validation Error", "Please select at least one day for weekly schedule.")
                return
        
        self.accept()
    
    def get_task_data(self) -> Dict[str, Any]:
        """Get edited task data."""
        schedule_type = self.schedule_type_combo.currentText().lower()
        time = self.time_edit.time().toString("HH:mm")
        
        schedule = {"type": schedule_type, "time": time}
        
        if schedule_type == "daily":
            interval = int(self.daily_interval_spin.currentText())
            if interval > 1:
                schedule["interval"] = interval
        elif schedule_type == "weekly":
            selected_days = [day for day, checkbox in self.day_checkboxes.items() if checkbox.isChecked()]
            schedule["days"] = selected_days
        elif schedule_type == "minute":
            schedule["interval"] = int(self.minute_interval_spin.currentText())
        
        return {
            "id": self.id_edit.text().strip(),
            "exe": self.exe_edit.text().strip(),
            "args": self.args_edit.text().strip(),
            "schedule": schedule,
        }


class ScheduledTasksPage(QWidget):
    """
    Page for managing scheduled tasks.
    
    Features:
    - Display tasks in table
    - Edit task configurations
    - Manual task execution
    - Initialize Windows Task Scheduler
    """
    
    def __init__(self, parent=None):
        """Initialize scheduled tasks page."""
        super().__init__(parent)
        self.tasks: List[Dict[str, Any]] = []
        self.current_worker: Optional[TaskExecutionWorker] = None
        self._setup_ui()
        self._load_tasks()
    
    def _setup_ui(self):
        """Setup page UI."""
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)
        
        # Header
        title = TitleLabel("Scheduled Tasks")
        root.addWidget(title)
        
        desc = BodyLabel("Manage scheduled task configurations and execute tasks manually.")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #808080;")
        root.addWidget(desc)
        
        # Scrollable content
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(16)
        
        # Tasks table card
        table_card = CardWidget(content)
        table_layout = QVBoxLayout()
        table_layout.setContentsMargins(16, 16, 16, 16)
        table_layout.setSpacing(12)
        
        # Table header with refresh button
        header_layout = QHBoxLayout()
        table_label = BodyLabel("Scheduled Tasks")
        table_label.setStyleSheet("font-weight: 600;")
        header_layout.addWidget(table_label)
        header_layout.addStretch()
        
        refresh_btn = PushButton("Refresh", content, FluentIcon.SYNC)
        refresh_btn.clicked.connect(self._load_tasks)
        header_layout.addWidget(refresh_btn)
        
        table_layout.addLayout(header_layout)
        
        # Tasks table
        self.tasks_table = QTableWidget()
        self.tasks_table.setColumnCount(3)
        self.tasks_table.setHorizontalHeaderLabels(["Task Name", "Schedule", "Actions"])
        self.tasks_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tasks_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.tasks_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        # Set minimum width for Actions column to ensure buttons are fully visible
        self.tasks_table.setColumnWidth(2, 180)
        self.tasks_table.verticalHeader().setDefaultSectionSize(50)
        self.tasks_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tasks_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table_layout.addWidget(self.tasks_table)
        
        table_card.setLayout(table_layout)
        content_layout.addWidget(table_card)
        
        # Log area card
        log_card = CardWidget(content)
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(16, 16, 16, 16)
        log_layout.setSpacing(12)
        
        log_label = BodyLabel("Execution Log")
        log_label.setStyleSheet("font-weight: 600;")
        log_layout.addWidget(log_label)
        
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMinimumHeight(200)
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
        log_layout.addWidget(self.log_viewer)
        
        log_actions = QHBoxLayout()
        log_actions.addStretch()
        clear_log_btn = PushButton("Clear Log", content)
        clear_log_btn.clicked.connect(self.log_viewer.clear)
        log_actions.addWidget(clear_log_btn)
        log_layout.addLayout(log_actions)
        
        log_card.setLayout(log_layout)
        content_layout.addWidget(log_card)
        
        content_layout.addStretch()
        scroll.setWidget(content)
        root.addWidget(scroll, stretch=1)
        
        # Action buttons
        action_layout = QHBoxLayout()
        action_layout.addStretch()
        
        self.init_btn = PrimaryPushButton("Initialize Scheduled Tasks", self, FluentIcon.SETTING)
        self.init_btn.clicked.connect(self._on_init_clicked)
        action_layout.addWidget(self.init_btn)
        
        root.addLayout(action_layout)
    
    def _load_tasks(self):
        """Load tasks from config file and populate table."""
        self.tasks = load_tasks()
        self._populate_table()
        self._append_log("INFO", f"Loaded {len(self.tasks)} scheduled task(s)")
    
    def _populate_table(self):
        """Populate table with tasks."""
        self.tasks_table.setRowCount(len(self.tasks))
        
        for row, task in enumerate(self.tasks):
            
            # Task name
            name_item = QTableWidgetItem(task.get("id", ""))
            self.tasks_table.setItem(row, 0, name_item)
            
            # Schedule info
            schedule = task.get("schedule", {})
            schedule_text = format_schedule_info(schedule)
            schedule_item = QTableWidgetItem(schedule_text)
            self.tasks_table.setItem(row, 1, schedule_item)
            
            # Actions
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(4, 4, 4, 4)
            actions_layout.setSpacing(8)
            
            edit_btn = PushButton("Edit", actions_widget, FluentIcon.EDIT)
            edit_btn.setMinimumWidth(80)
            edit_btn.setMinimumHeight(30)
            edit_btn.clicked.connect(lambda checked, r=row: self._on_edit_clicked(r))
            actions_layout.addWidget(edit_btn)
            
            execute_btn = PushButton("Execute", actions_widget, FluentIcon.PLAY)
            execute_btn.setMinimumWidth(80)
            execute_btn.setMinimumHeight(30)
            execute_btn.clicked.connect(lambda checked, r=row: self._on_execute_clicked(r))
            actions_layout.addWidget(execute_btn)
            
            actions_layout.addStretch()
            self.tasks_table.setCellWidget(row, 2, actions_widget)
    
    def _on_edit_clicked(self, row: int):
        """Handle edit button click."""
        if row < 0 or row >= len(self.tasks):
            return
        
        task = self.tasks[row]
        dialog = TaskEditDialog(task, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            edited_task = dialog.get_task_data()
            self.tasks[row] = edited_task
            
            if save_tasks(self.tasks):
                self._append_log("SUCCESS", f"Task '{edited_task.get('id')}' updated successfully")
                self._load_tasks()
            else:
                self._append_log("ERROR", "Failed to save task configuration")
                QMessageBox.warning(self, "Error", "Failed to save task configuration")
    
    def _on_execute_clicked(self, row: int):
        """Handle execute button click."""
        if row < 0 or row >= len(self.tasks):
            return
        
        if self.current_worker and self.current_worker.isRunning():
            self._append_log("WARNING", "Task is already running. Please wait for completion.")
            return
        
        task = self.tasks[row]
        task_id = task.get("id", "Unknown")
        
        self._append_log("INFO", f"Starting manual execution of task: {task_id}")
        
        # Create and start worker
        self.current_worker = TaskExecutionWorker(task)
        self.current_worker.log_signal.connect(self._on_log_received)
        self.current_worker.finished.connect(self._on_worker_finished)
        self.current_worker.start()
        
        # Disable execute buttons
        self._set_execution_enabled(False)
    
    def _on_init_clicked(self):
        """Handle initialize scheduled tasks button click."""
        reply = QMessageBox.question(
            self,
            "Initialize Scheduled Tasks",
            "This will register all scheduled tasks to Windows Task Scheduler. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self._append_log("INFO", "Initializing scheduled tasks...")
            try:
                schedule_task_module.main()
                self._append_log("SUCCESS", "Scheduled tasks initialized successfully")
                QMessageBox.information(self, "Success", "Scheduled tasks initialized successfully")
            except Exception as e:
                self._append_log("ERROR", f"Failed to initialize scheduled tasks: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to initialize scheduled tasks:\n{str(e)}")
    
    @Slot(str, str)
    def _on_log_received(self, level: str, message: str):
        """Handle log signal from worker."""
        self._append_log(level, message)
    
    @Slot(dict)
    def _on_worker_finished(self, result: Dict[str, Any]):
        """Handle worker finished signal."""
        self._set_execution_enabled(True)
        
        if result.get("success"):
            self._append_log("SUCCESS", "Task execution completed successfully")
        else:
            error_msg = result.get("message", "Unknown error")
            self._append_log("ERROR", f"Task execution failed: {error_msg}")
    
    def _set_execution_enabled(self, enabled: bool):
        """Set execution buttons enabled state."""
        for row in range(self.tasks_table.rowCount()):
            actions_widget = self.tasks_table.cellWidget(row, 2)
            if actions_widget:
                for child in actions_widget.findChildren(QPushButton):
                    if child.text() == "Execute":
                        child.setEnabled(enabled)
        self.init_btn.setEnabled(enabled)
    
    def _append_log(self, level: str, message: str):
        """Append log message to log viewer."""
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
