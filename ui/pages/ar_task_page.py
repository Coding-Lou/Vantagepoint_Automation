"""
AR (Accounts Receivable) task page with modern Fluent Design.

Task: Generate AR statements with email configuration.
"""
from typing import Dict, Any, Tuple, Optional

from PySide6.QtWidgets import (
    QLabel,
    QFormLayout,
    QHBoxLayout,
)
from PySide6.QtCore import QDate, Qt

from qfluentwidgets import (
    DatePicker,
    LineEdit,
    BodyLabel,
    TextEdit,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.ar_worker import ARWorker
import tools.util as util_module
import tools.config_manager as config_manager


class ARTaskPage(BaseTaskPage):
    """
    Task page for AR statement generation.
    """
    
    def __init__(self, parent=None):
        """Initialize AR task page."""
        super().__init__(
            "ar",
            "AR Notification",
            parent
        )
    
    def _get_page_description(self) -> str:
        """Get page description."""
        return "Generate and send AR statements to clients via email."
    
    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)
        
        tutorial_text = BodyLabel(
            "1. Select the statement date using the calendar picker.\n"
            "2. Configure email settings (From, CC, Subject, Body) by clicking 'Edit' button.\n"
            "3. Click 'Save' to save your email configuration.\n"
            "4. Click 'Execute' to generate and send AR statement emails."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Statement date section
        date_label = BodyLabel("Statement Date")
        date_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(date_label)
        
        self.date_edit = DatePicker()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(lambda date: print(date.toString()))
        # Only allow selecting date from calendar popup (no manual typing)
        try:
            line = self.date_edit.lineEdit()
            if line:
                line.setReadOnly(True)
        except Exception:
            pass
        self.date_edit.setStyleSheet(
            f"""
            QDateEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )
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
        self.mail_subject_edit.setPlaceholderText("AR Statement")
        self.mail_subject_edit.setReadOnly(True)
        email_form.addRow("Subject:", self.mail_subject_edit)
        
        body_label = BodyLabel("Body (HTML):")
        body_label.setStyleSheet("font-size: 12px;")
        self.config_layout.addLayout(email_form)
        self.config_layout.addWidget(body_label)
        
        self.mail_body_edit = TextEdit()
        self.mail_body_edit.setAcceptRichText(True)
        self.mail_body_edit.setMaximumHeight(500)
        self.mail_body_edit.setPlaceholderText("Enter email body HTML content...")
        self.mail_body_edit.setReadOnly(True)
        self.mail_body_edit.setStyleSheet(
            f"""
            QTextEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )
        self.config_layout.addWidget(self.mail_body_edit)
        
        # Track edit mode
        self.email_edit_mode = False
    
    def _load_config(self) -> None:
        """Load configuration from config file."""
        # Email configuration
        self.mail_from_edit.setText(util_module.get_config(["AR", "FROM"]) or "")
        self.mail_cc_edit.setText(util_module.get_config(["AR", "CC"]) or "")
        self.mail_subject_edit.setText(util_module.get_config(["AR", "SUBJECT"]) or "")
        self.mail_body_edit.setHtml(util_module.get_config(["AR", "BODY"]) or "")
    
    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        if not self.date_edit.date().isValid():
            return False, "Please select a valid statement date"
        return True, ""
    
    def _get_params(self) -> Dict[str, Any]:
        """Get task parameters from UI."""
        date = self.date_edit.date().toString("yyyy-MM-dd")
        
        return {
            "statement_date": date,
            "mail_from": self.mail_from_edit.text() or None,
            "mail_cc": self.mail_cc_edit.text() or None,
            "mail_subject": self.mail_subject_edit.text() or None,
            "mail_body": self.mail_body_edit.toHtml() or None,
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> ARWorker:
        """Create AR worker."""
        return ARWorker(params)
    
    def _on_edit_email_clicked(self) -> None:
        """Handle Edit button click - enable editing of email fields."""
        if not self.email_edit_mode:
            # Enable editing
            self.mail_from_edit.setReadOnly(False)
            self.mail_cc_edit.setReadOnly(False)
            self.mail_subject_edit.setReadOnly(False)
            self.mail_body_edit.setReadOnly(False)
            
            # Switch to plain text mode to show HTML tags
            # Get the original HTML source from config to display as plain text
            html_source = util_module.get_config(["AR", "BODY"]) or ""
            self.mail_body_edit.setPlainText(html_source)
            
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
            
            # In edit mode, get the plain text (HTML source) that user edited
            html_source = self.mail_body_edit.toPlainText()
            mail_body = html_source
            
            # Save to config using config_manager
            updates = {
                ("AR", "FROM"): mail_from,
                ("AR", "CC"): mail_cc,
                ("AR", "SUBJECT"): mail_subject,
                ("AR", "BODY"): mail_body,
            }
            
            success = config_manager.update_multiple_config_values(updates)
            
            if success:
                # Disable editing
                self.mail_from_edit.setReadOnly(True)
                self.mail_cc_edit.setReadOnly(True)
                self.mail_subject_edit.setReadOnly(True)
                self.mail_body_edit.setReadOnly(True)
                
                # Switch back to HTML rendering mode (hide HTML tags)
                # Set the HTML content to display rendered version
                self.mail_body_edit.setHtml(mail_body)
                
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
