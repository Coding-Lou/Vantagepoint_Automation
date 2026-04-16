"""
Revenue Accrual task page with modern Fluent Design.

Task: Generate revenue accrual workbooks.
"""
from typing import Dict, Any, Tuple, Optional

from PySide6.QtWidgets import (
    QVBoxLayout,
    QHBoxLayout,
)
from PySide6.QtCore import Qt

from qfluentwidgets import (
    ComboBox,
    LineEdit,
    BodyLabel,
    PrimaryPushButton,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.revenue_accrual_worker import RevenueAccrualWorker
from workers.base_worker import BaseWorker
from ui.services.app_context import get_app_context
import tools.util as util_module
import core.project_status as ps_module
import core.revenue_accrual as revenue_accrual_module


class UpdateRevGenWorker(BaseWorker):
    """Worker for executing update_revgen in background thread."""
    
    def __init__(self):
        """Initialize Update Rev-Gen worker."""
        super().__init__("revenue_accrual.update_revgen", {}, auto_login_enabled=False)
    
    def execute(self) -> Dict[str, Any]:
        """Execute update_revgen method."""
        try:
            self.log_signal.emit("INFO", "Starting Rev-Gen update...")
            # Execute the update_revgen function
            revenue_accrual_module.update_revgen()
            self.log_signal.emit("SUCCESS", "Rev-Gen update completed successfully")
            return {
                "success": True,
                "message": "Rev-Gen update completed successfully"
            }
        except Exception as e:
            error_msg = str(e)
            self.log_signal.emit("ERROR", f"Rev-Gen update failed: {error_msg}")
            return {
                "success": False,
                "message": f"Rev-Gen update failed: {error_msg}"
            }


class RevenueAccrualTaskPage(BaseTaskPage):
    """
    Task page for revenue accrual generation.
    """

    def __init__(self, parent=None):
        """Initialize Revenue Accrual task page."""
        # Initialize period_data_map before calling super().__init__()
        self.period_data_map: Dict[str, str] = {}

        super().__init__(
            "revenue_accrual",
            "Revenue Accrual",
            parent,
        )

        # Connect to login state changes to reload periods when user logs in
        app_context = get_app_context()
        app_context.login_state_changed.connect(self._on_login_state_changed)
        
        # Initialize update revgen worker
        self.update_revgen_worker: Optional[UpdateRevGenWorker] = None
        
        # Add Update Rev-Gen button above Execution Log
        self._add_update_revgen_button()

    def _get_page_description(self) -> str:
        """Get page description."""
        return (
            "Generate earned revenue accrual workbooks by combining invoices, "
            "general ledger data, budgets, and labor information."
        )

    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)

        tutorial_text = BodyLabel(
            "1. Select an accounting period from the dropdown menu.\n"
            "2. Click 'Execute' to generate the revenue accrual workbooks for the selected period.\n"
            "3. Review the generated Excel files in the 'revenue accural' folder."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)

        # Spacer
        self.config_layout.addSpacing(12)

        # Period section
        period_label = BodyLabel("Accounting Period")
        period_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(period_label)

        self.period_combo = ComboBox()
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        arrow_color = ThemeColors.text_secondary()
        self.period_combo.setStyleSheet(
            f"""
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
        """
        )

        # Load periods from project_status.print_period()
        self._load_periods()
        self.config_layout.addWidget(self.period_combo)

        # Spacer at the bottom of the config area
        self.config_layout.addSpacing(12)

    def _load_config(self) -> None:
        """Load configuration from config file (if any)."""
        # Revenue accrual does not have saved config for now.
        pass

    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        """Handle login state change - reload periods when user logs in."""
        if is_logged_in:
            self._load_periods()

    def _load_periods(self) -> None:
        """Load periods from project_status.print_period() and populate the dropdown."""
        # Ensure period_combo exists
        if not hasattr(self, "period_combo"):
            return

        try:
            # Update headers to ensure fresh authentication for period API
            ps_module.HEADERS = util_module.set_headers()

            # Get period data from project_status.print_period()
            period_data = ps_module.print_period()

            # Clear existing items and mapping
            self.period_combo.clear()
            self.period_data_map.clear()

            # Populate dropdown with formatted display text
            for p in period_data:
                display_text = (
                    f"{p['Period']} | From: {p['AccountPdStart'][:10]} To: {p['AccountPdEnd'][:10]}"
                )
                period_value = str(p["Period"])

                # Store mapping: display text -> Period value
                self.period_data_map[display_text] = period_value

                # Add to dropdown
                self.period_combo.addItem(display_text)

            # Select first item if available
            if self.period_combo.count() > 0:
                self.period_combo.setCurrentIndex(0)

        except Exception as e:
            # If loading fails, add an error message
            try:
                self.period_combo.clear()
                self.period_combo.addItem(
                    "Error loading periods - please check connection"
                )
            except Exception:
                pass
            print(f"Error loading periods: {e}")

    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        # Check if period is selected
        if self.period_combo.currentIndex() < 0 or self.period_combo.currentText() == "":
            return False, "Please select an accounting period"

        # Check if period data is valid (not error message)
        current_text = self.period_combo.currentText()
        if current_text.startswith("Error loading periods"):
            return False, "Please ensure periods are loaded correctly"

        # Verify period value exists in mapping
        if current_text not in self.period_data_map:
            return False, "Invalid period selection"

        return True, ""

    def _get_params(self) -> Dict[str, Any]:
        """Get task parameters from UI."""
        # Get Period value from selected display text
        current_text = self.period_combo.currentText()
        period = self.period_data_map.get(current_text, "")

        # Ensure period is a string
        if period:
            period = str(period)

        return {
            "period": period,
        }

    def _create_worker(self, params: Dict[str, Any]) -> RevenueAccrualWorker:
        """Create Revenue Accrual worker."""
        return RevenueAccrualWorker(params)
    
    def _add_update_revgen_button(self) -> None:
        """Add Update Rev-Gen button above Execution Log section."""
        # Find the log_card in the content_layout and insert button before it
        content_layout = self._content_widget.layout()
        if not content_layout:
            return
        
        # Find the index of log_card in the layout
        log_card_index = -1
        for i in range(content_layout.count()):
            item = content_layout.itemAt(i)
            if item and item.widget() == self.log_card:
                log_card_index = i
                break
        
        if log_card_index < 0:
            return
        
        # Create button card
        from qfluentwidgets import CardWidget
        button_card = CardWidget(self._content_widget)
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(16, 16, 16, 16)
        button_layout.setSpacing(12)
        
        # Create Update Rev-Gen button
        self.update_revgen_btn = PrimaryPushButton("Update Rev-Gen", button_card)
        self.update_revgen_btn.setMinimumWidth(120)
        self.update_revgen_btn.setMinimumHeight(36)
        self.update_revgen_btn.clicked.connect(self._on_update_revgen_clicked)
        
        button_layout.addWidget(self.update_revgen_btn)
        button_layout.addStretch()
        
        button_card.setLayout(button_layout)
        
        # Insert button card before log_card
        content_layout.insertWidget(log_card_index, button_card)
    
    def _on_update_revgen_clicked(self) -> None:
        """Handle Update Rev-Gen button click."""
        # Check login status before running
        try:
            is_logged_in = bool(util_module.check_login())
        except Exception:
            is_logged_in = False

        if not is_logged_in:
            self._append_log("WARNING", "Login required. Opening sign-in dialog...")
            # Try to trigger login dialog on parent MainWindow
            parent = self.parent()
            while parent:
                if hasattr(parent, "_on_login_clicked"):
                    try:
                        parent._on_login_clicked()
                    except Exception:
                        pass
                    break
                parent = parent.parent()
            return

        # Check if worker is already running
        if self.update_revgen_worker and self.update_revgen_worker.isRunning():
            self._append_log("WARNING", "Rev-Gen update is already running. Please wait for completion.")
            return

        # Clean up any existing worker
        if self.update_revgen_worker:
            try:
                self.update_revgen_worker.log_signal.disconnect()
            except Exception:
                pass
            try:
                self.update_revgen_worker.finished.disconnect()
            except Exception:
                pass
            self.update_revgen_worker = None

        # Create and start worker
        self.update_revgen_worker = UpdateRevGenWorker()
        self.update_revgen_worker.log_signal.connect(self._on_log_received)
        self.update_revgen_worker.finished.connect(self._on_update_revgen_finished)
        
        # Update UI state
        self.update_revgen_btn.setEnabled(False)
        self.log_viewer.clear()
        self._append_log("INFO", "Starting Rev-Gen update...")
        
        # Start worker in background thread
        self.update_revgen_worker.start()
    
    def _on_update_revgen_finished(self, result: Dict[str, Any]) -> None:
        """Handle update revgen worker finished signal."""
        # Disconnect signals
        if self.update_revgen_worker:
            try:
                self.update_revgen_worker.log_signal.disconnect()
            except Exception:
                pass
            try:
                self.update_revgen_worker.finished.disconnect()
            except Exception:
                pass
            self.update_revgen_worker = None
        
        # Update UI state
        self.update_revgen_btn.setEnabled(True)
        
        if result.get("success"):
            self._append_log("SUCCESS", "Rev-Gen update completed successfully")
        else:
            error_msg = result.get("message", "Unknown error")
            self._append_log("ERROR", f"Rev-Gen update failed: {error_msg}")

