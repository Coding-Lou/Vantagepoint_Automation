"""
AP (Accounts Payable) task page with modern Fluent Design.

Task: Generate AP remittance advice with email configuration.
"""
from typing import Dict, Any, Tuple, Optional

from PySide6.QtWidgets import (
    QLineEdit,
    QTextEdit,
    QLabel,
    QFormLayout,
    QHBoxLayout,
    QMessageBox,
)
from PySide6.QtCore import QDate, Qt, Slot

from qfluentwidgets import (
    DatePicker,
    LineEdit,
    BodyLabel,
    CardWidget,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
    isDarkTheme,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.ap_worker import APWorker
import tools.util as util_module
import tools.config_manager as config_manager
import json


class APTaskPage(BaseTaskPage):
    """
    Task page for AP remittance generation.
    """
    
    def __init__(self, parent=None):
        """Initialize AP task page."""
        super().__init__(
            "ap",
            "AP Remittance",
            parent
        )
    
    def _get_page_description(self) -> str:
        """Get page description."""
        return "Generate and send AP remittance advice to vendors via email."
    
    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)
        
        tutorial_text = BodyLabel(
            "1. Select the remittance date using the calendar picker.\n"
            "2. Configure email settings (From, CC, Subject, Body) by clicking 'Edit' button.\n"
            "3. Click 'Save' to save your email configuration.\n"
            "4. Click 'Execute' to generate and send remittance advice emails."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Remittance date section
        date_label = BodyLabel("Remittance Date")
        date_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(date_label)
        
        self.date_edit = DatePicker()
        self.date_edit.setDate(QDate.currentDate())
        #region agent log
        try:
            log_path = r"c:\cursor\.cursor\debug.log"
            date_attr = getattr(self.date_edit, 'date', None)
            date_type = type(date_attr).__name__ if date_attr is not None else 'None'
            is_callable = callable(date_attr)
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    "sessionId": "debug-session",
                    "runId": "run1",
                    "hypothesisId": "C",
                    "location": "ap_task_page.py:82",
                    "message": "After DatePicker initialization and setDate",
                    "data": {
                        "date_type": date_type,
                        "is_callable": is_callable,
                        "date_repr": str(date_attr)[:100] if date_attr is not None else None
                    },
                    "timestamp": int(__import__('time').time() * 1000)
                }) + '\n')
        except Exception as e:
            pass
        #endregion
        self.date_edit.dateChanged.connect(lambda date: print(date.toString()))
        # Only allow selecting date from calendar popup (no manual typing)
        # NOTE: Don't call QDateEdit.setReadOnly(True) here; in some Qt builds it
        # can also block changes via the popup calendar. We only lock the line edit.
        try:
            line = self.date_edit.lineEdit()
            if line:
                line.setReadOnly(True)
        except Exception:
            pass
        # Apply theme-aware style to date picker
        self._apply_date_edit_style()
        self.config_layout.addWidget(self.date_edit)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Email configuration section with Edit/Save buttons
        email_header_layout = QHBoxLayout()
        email_label = BodyLabel("Email Configuration")
        email_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        email_header_layout.addWidget(email_label)
        email_header_layout.addStretch()
        
        # Edit/Save buttons
        self.email_edit_btn = PushButton("Edit", self, FluentIcon.EDIT)
        self.email_edit_btn.setMaximumWidth(80)
        self.email_edit_btn.clicked.connect(self._on_edit_email_clicked)
        email_header_layout.addWidget(self.email_edit_btn)
        
        self.email_save_btn = PrimaryPushButton("Save", self, FluentIcon.SAVE)
        self.email_save_btn.setMaximumWidth(80)
        self.email_save_btn.setEnabled(False)
        self.email_save_btn.clicked.connect(self._on_save_email_clicked)
        email_header_layout.addWidget(self.email_save_btn)
        
        self.config_layout.addLayout(email_header_layout)
        
        # Email form
        email_form = QFormLayout()
        email_form.setSpacing(12)
        email_form.setContentsMargins(0, 0, 0, 0)
        
        self.mail_from_edit = LineEdit()
        self.mail_from_edit.setPlaceholderText("sender@example.com")
        self.mail_from_edit.setReadOnly(True)
        email_form.addRow("From:", self.mail_from_edit)
        
        self.mail_cc_edit = LineEdit()
        self.mail_cc_edit.setPlaceholderText("cc@example.com")
        self.mail_cc_edit.setReadOnly(True)
        email_form.addRow("CC:", self.mail_cc_edit)
        
        self.mail_subject_edit = LineEdit()
        self.mail_subject_edit.setPlaceholderText("Remittance Advice")
        self.mail_subject_edit.setReadOnly(True)
        email_form.addRow("Subject:", self.mail_subject_edit)
        
        body_label = BodyLabel("Body (HTML):")
        body_label.setStyleSheet("font-size: 12px;")
        self.config_layout.addLayout(email_form)
        self.config_layout.addWidget(body_label)
        
        self.mail_body_edit = QTextEdit()
        self.mail_body_edit.setAcceptRichText(True)
        self.mail_body_edit.setMaximumHeight(100)
        self.mail_body_edit.setPlaceholderText("Enter email body HTML content...")
        self.mail_body_edit.setReadOnly(True)
        # Apply theme-aware style to mail body editor
        self._apply_mail_body_style()
        self.config_layout.addWidget(self.mail_body_edit)
        
        # Track edit mode
        self.email_edit_mode = False
    
    def _apply_date_edit_style(self) -> None:
        """Apply theme-aware style to date picker."""
        self.date_edit.setStyleSheet(
            f"""
            QDateEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )

    def _apply_mail_body_style(self) -> None:
        """Apply theme-aware style to mail body editor."""
        self.mail_body_edit.setStyleSheet(
            f"""
            QTextEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )
    
    def _load_config(self) -> None:
        """Load configuration from config file."""
        # Email configuration
        self.mail_from_edit.setText(util_module.get_config(["AP", "FROM"]) or "")
        self.mail_cc_edit.setText(util_module.get_config(["AP", "CC"]) or "")
        self.mail_subject_edit.setText(util_module.get_config(["AP", "SUBJECT"]) or "")
        self.mail_body_edit.setHtml(util_module.get_config(["AP", "BODY"]) or "")
    
    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        #region agent log
        try:
            log_path = r"c:\cursor\.cursor\debug.log"
            date_attr = getattr(self.date_edit, 'date', None)
            date_type = type(date_attr).__name__ if date_attr is not None else 'None'
            is_callable = callable(date_attr)
            is_valid = date_attr.isValid() if date_attr is not None else False
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    "sessionId": "debug-session",
                    "runId": "post-fix",
                    "hypothesisId": "A",
                    "location": "ap_task_page.py:178",
                    "message": "Checking date_edit.date attribute (post-fix)",
                    "data": {
                        "date_type": date_type,
                        "is_callable": is_callable,
                        "is_valid": is_valid,
                        "date_repr": str(date_attr)[:100] if date_attr is not None else None
                    },
                    "timestamp": int(__import__('time').time() * 1000)
                }) + '\n')
        except Exception as e:
            pass
        #endregion
        if not self.date_edit.date.isValid():
            return False, "Please select a valid remittance date"
        return True, ""
    
    def _get_params(self) -> Dict[str, Any]:
        """Get task parameters from UI."""
        #region agent log
        try:
            log_path = r"c:\cursor\.cursor\debug.log"
            date_attr = getattr(self.date_edit, 'date', None)
            date_type = type(date_attr).__name__ if date_attr is not None else 'None'
            is_callable = callable(date_attr)
            date_str = date_attr.toString("yyyy-MM-dd") if date_attr is not None else None
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    "sessionId": "debug-session",
                    "runId": "post-fix",
                    "hypothesisId": "B",
                    "location": "ap_task_page.py:184",
                    "message": "Getting date in _get_params (post-fix)",
                    "data": {
                        "date_type": date_type,
                        "is_callable": is_callable,
                        "date_str": date_str,
                        "date_repr": str(date_attr)[:100] if date_attr is not None else None
                    },
                    "timestamp": int(__import__('time').time() * 1000)
                }) + '\n')
        except Exception as e:
            pass
        #endregion
        date = self.date_edit.date.toString("yyyy-MM-dd")
        
        return {
            "remittance_date": date,
            "exclude_vendors": None,  # Exclude vendors feature removed
            "mail_from": self.mail_from_edit.text() or None,
            "mail_cc": self.mail_cc_edit.text() or None,
            "mail_subject": self.mail_subject_edit.text() or None,
            "mail_body": self.mail_body_edit.toHtml() or None,
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> APWorker:
        """Create AP worker."""
        return APWorker(params)
    
    @Slot(dict)
    def _on_worker_finished(self, result: Dict[str, Any]) -> None:
        """
        Handle worker finished signal with completion dialog.
        
        Args:
            result: Result dictionary
        """
        # Call parent implementation first
        super()._on_worker_finished(result)
        
        # Show completion dialog if task succeeded
        if result.get("success"):
            QMessageBox.information(
                self,
                "Task Completed",
                "AP Remittance task has been completed successfully."
            )
    
    def _on_edit_email_clicked(self) -> None:
        """Handle Edit button click - enable editing of email fields."""
        if not self.email_edit_mode:
            # Enable editing
            self.mail_from_edit.setReadOnly(False)
            self.mail_cc_edit.setReadOnly(False)
            self.mail_subject_edit.setReadOnly(False)
            self.mail_body_edit.setReadOnly(False)
            
            # Update button states
            self.email_edit_btn.setEnabled(False)
            self.email_save_btn.setEnabled(True)
            
            self.email_edit_mode = True
    
    def _on_save_email_clicked(self) -> None:
        """Handle Save button click - save email configuration to config.json."""
        try:
            # Get current values from UI
            mail_from = self.mail_from_edit.text()
            mail_cc = self.mail_cc_edit.text()
            mail_subject = self.mail_subject_edit.text()
            mail_body = self.mail_body_edit.toHtml()
            
            # Save to config using config_manager
            updates = {
                ("AP", "FROM"): mail_from,
                ("AP", "CC"): mail_cc,
                ("AP", "SUBJECT"): mail_subject,
                ("AP", "BODY"): mail_body,
            }
            
            success = config_manager.update_multiple_config_values(updates)
            
            if success:
                # Disable editing
                self.mail_from_edit.setReadOnly(True)
                self.mail_cc_edit.setReadOnly(True)
                self.mail_subject_edit.setReadOnly(True)
                self.mail_body_edit.setReadOnly(True)
                
                # Update button states
                self.email_edit_btn.setEnabled(True)
                self.email_save_btn.setEnabled(False)
                
                self.email_edit_mode = False
                
                # Show success message in log
                self._append_log("SUCCESS", "Email configuration saved successfully")
            else:
                self._append_log("ERROR", "Failed to save email configuration")
        except Exception as e:
            self._append_log("ERROR", f"Error saving email configuration: {str(e)}")