"""
Statement Check task page with modern Fluent Design.

Task: Check statement status for invoices.
"""
from typing import Dict, Any, Tuple, List

from PySide6.QtWidgets import (
    QHBoxLayout,
)
from qfluentwidgets import (
    ComboBox,
    LineEdit,
    BodyLabel,
    PushButton,
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
        super().__init__(
            "statement_check",
            "Vendor Statement Check",
            parent
        )

    def _get_page_description(self) -> str:
        return "Check statement status for vendor invoices to verify if they have been vouched."

    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)

        tutorial_text = BodyLabel(
            "1. Type part of the vendor name and click 'Search'.\n"
            "2. Select the correct vendor from the results list.\n"
            "3. Enter invoice numbers separated by commas (e.g., INV001, INV002).\n"
            "4. Click 'Execute' to check the statement status."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)

        self.config_layout.addSpacing(12)

        # Vendor search section
        vendor_label = BodyLabel("Vendor")
        vendor_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(vendor_label)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)

        self.vendor_search_edit = LineEdit()
        self.vendor_search_edit.setPlaceholderText("Type vendor name...")
        self.vendor_search_edit.setStyleSheet(
            f"LineEdit {{ {ThemeColors.get_input_style()} }}"
        )
        self.vendor_search_edit.returnPressed.connect(self._on_search)
        search_row.addWidget(self.vendor_search_edit, 1)

        self.search_btn = PushButton("Search")
        self.search_btn.clicked.connect(self._on_search)
        search_row.addWidget(self.search_btn)

        self.config_layout.addLayout(search_row)

        # Results combo (hidden until search returns results)
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        arrow_color = ThemeColors.text_secondary()
        self.vendor_combo = ComboBox()
        self.vendor_combo.setStyleSheet(f"""
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
        self.vendor_combo.setVisible(False)
        # internal map: display name -> vendor key
        self._vendor_key_map: dict[str, str] = {}
        self.config_layout.addWidget(self.vendor_combo)

        self.config_layout.addSpacing(12)

        # Invoice numbers section
        invoice_label = BodyLabel("Invoice Numbers")
        invoice_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(invoice_label)

        self.invoice_edit = LineEdit()
        self.invoice_edit.setPlaceholderText("INV001, INV002, INV003 (comma-separated)")
        self.invoice_edit.setStyleSheet(
            f"LineEdit {{ {ThemeColors.get_input_style()} }}"
        )
        self.config_layout.addWidget(self.invoice_edit)

    def _on_search(self) -> None:
        """Call util.get_firm_key() with the typed query and populate the combo."""
        query = self.vendor_search_edit.text().strip()
        if not query:
            return

        self.search_btn.setEnabled(False)
        self.search_btn.setText("Searching...")

        try:
            headers = util_module.set_headers()
            results: List[Dict] = util_module.get_firm_key(query=query, headers=headers) or []
        except Exception:
            results = []
        finally:
            self.search_btn.setEnabled(True)
            self.search_btn.setText("Search")

        self._vendor_key_map.clear()
        self.vendor_combo.clear()

        if not results:
            self.vendor_combo.setVisible(False)
            return

        for item in results:
            name = item.get("name", "")
            key = item.get("key", "")
            if name:
                self._vendor_key_map[name] = key
                self.vendor_combo.addItem(name)

        self.vendor_combo.setCurrentIndex(0)
        self.vendor_combo.setVisible(True)

    def _load_config(self) -> None:
        """Reload config (no-op here; search is on-demand)."""
        pass

    def _validate_params(self) -> Tuple[bool, str]:
        if not self.vendor_combo.isVisible() or self.vendor_combo.count() == 0:
            return False, "Please search for and select a vendor first"

        vendor_key = self._vendor_key_map.get(self.vendor_combo.currentText(), "")
        if not vendor_key:
            return False, "Invalid vendor selection"

        invoice_numbers = self.invoice_edit.text().strip()
        if not invoice_numbers:
            return False, "Please enter at least one invoice number"

        invoice_list = [inv.strip() for inv in invoice_numbers.split(",") if inv.strip()]
        if not invoice_list:
            return False, "Please enter at least one valid invoice number"

        return True, ""

    def _get_params(self) -> Dict[str, Any]:
        vendor_key = self._vendor_key_map.get(self.vendor_combo.currentText(), "")
        invoice_numbers = self.invoice_edit.text().strip()
        return {
            "vendor_key": vendor_key,
            "invoice_numbers": invoice_numbers,
        }

    def _create_worker(self, params: Dict[str, Any]) -> StatementCheckWorker:
        return StatementCheckWorker(params)
