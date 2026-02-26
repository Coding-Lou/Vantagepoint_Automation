"""
Statement Check task page with modern Fluent Design.

Task: Check statement status for invoices.
"""
from typing import Dict, Any, Tuple

from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
)
from PySide6.QtCore import Qt

from qfluentwidgets import (
    ComboBox,
    LineEdit,
    BodyLabel,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.statement_check_worker import StatementCheckWorker
import tools.util as util_module


class StatementCheckTaskPage(BaseTaskPage):
    """
    Task page for statement check.
    """
    
    def __init__(self, parent=None):
        """Initialize Statement Check task page."""
        super().__init__(
            "statement_check",
            "Vendor Statement Check",
            parent
        )
    
    def _get_page_description(self) -> str:
        """Get page description."""
        return "Check statement status for vendor invoices to verify if they have been vouched."
    
    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)
        
        tutorial_text = BodyLabel(
            "1. Select a client from the dropdown menu.\n"
            "2. Enter invoice numbers separated by commas (e.g., INV001, INV002, INV003).\n"
            "3. Click 'Execute' to check the statement status for the invoices."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Client selection section
        client_label = BodyLabel("Client")
        client_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(client_label)
        
        self.client_combo = ComboBox()
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        arrow_color = ThemeColors.text_secondary()
        self.client_combo.setStyleSheet(f"""
            ComboBox {{
                padding: 6px;
                border: 1px solid {border};
                border-radius: 4px;
                background-color: {bg};
                color: {text};
            }}
            ComboBox::drop-down {{
                border: none;
                padding-right: 8px;
            }}
            ComboBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid {arrow_color};
                margin-right: 4px;
            }}
        """)
        
        # Load clients from config
        self._load_clients()
        self.config_layout.addWidget(self.client_combo)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Invoice numbers section
        invoice_label = BodyLabel("Invoice Numbers")
        invoice_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(invoice_label)
        
        self.invoice_edit = LineEdit()
        self.invoice_edit.setPlaceholderText("INV001, INV002, INV003 (comma-separated)")
        self.invoice_edit.setStyleSheet(
            f"""
            LineEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )
        self.config_layout.addWidget(self.invoice_edit)
    
    def _load_clients(self) -> None:
        """Load clients from config.json STATEMENT_CHECK array."""
        try:
            statement_check_config = util_module.get_config(["STATEMENT_CHECK"])
            if not statement_check_config:
                return
            
            # Clear existing items
            self.client_combo.clear()
            
            # Add clients from config
            # STATEMENT_CHECK is an array of objects like [{"ANIXTER CAD": "ANICAN"}, ...]
            for client_obj in statement_check_config:
                if isinstance(client_obj, dict):
                    # Get the key (display name) and value (vendor key)
                    for display_name, vendor_key in client_obj.items():
                        # Store vendor_key as item data
                        self.client_combo.addItem(display_name, vendor_key)
            
            # Set default selection to first item if available
            if self.client_combo.count() > 0:
                self.client_combo.setCurrentIndex(0)
        except Exception as e:
            # If loading fails, add a placeholder
            self.client_combo.addItem("No clients available", "")
    
    def _load_config(self) -> None:
        """Load configuration from config file (if any)."""
        # Reload clients in case config was updated
        self._load_clients()
    
    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        # Check if a client is selected
        if self.client_combo.currentIndex() < 0:
            return False, "Please select a client"
        
        vendor_key = self.client_combo.currentData()
        if not vendor_key:
            return False, "Invalid client selection"
        
        # Check if invoice numbers are provided
        invoice_numbers = self.invoice_edit.text().strip()
        if not invoice_numbers:
            return False, "Please enter at least one invoice number"
        
        # Validate that there's at least one non-empty invoice number
        invoice_list = [inv.strip() for inv in invoice_numbers.split(",") if inv.strip()]
        if not invoice_list:
            return False, "Please enter at least one valid invoice number"
        
        return True, ""
    
    def _get_params(self) -> Dict[str, Any]:
        """Get task parameters from UI."""
        # Get vendor_key from selected client's item data
        vendor_key = self.client_combo.currentData()
        invoice_numbers = self.invoice_edit.text().strip()
        
        return {
            "vendor_key": vendor_key,
            "invoice_numbers": invoice_numbers,  # Comma-separated string
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> StatementCheckWorker:
        """Create Statement Check worker."""
        return StatementCheckWorker(params)
