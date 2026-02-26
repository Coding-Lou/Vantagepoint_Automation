"""
Revenue Accrual task page with modern Fluent Design.

Task: Generate revenue accrual workbooks.
"""
from typing import Dict, Any, Tuple

from PySide6.QtWidgets import (
    QVBoxLayout,
)
from PySide6.QtCore import Qt

from qfluentwidgets import (
    ComboBox,
    LineEdit,
    BodyLabel,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.revenue_accrual_worker import RevenueAccrualWorker
from ui.services.app_context import get_app_context
import tools.util as util_module
import core.project_status as ps_module


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

