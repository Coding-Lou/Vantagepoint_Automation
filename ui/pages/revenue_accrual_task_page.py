"""
Revenue Accrual task page with Fluent Design.

Layout (top -> bottom inside config card):
  1. SOP — compact usage instructions
  2. Accounting Period — dropdown
  3. Project Scope — dropdown | Total count | Download List / Add New note
"""
from typing import Dict, Any, Tuple, Optional

from PySide6.QtWidgets import (
    QVBoxLayout,
    QHBoxLayout,
    QStackedWidget,
    QWidget,
)
from PySide6.QtCore import Qt, Signal, Slot, QThread

from qfluentwidgets import (
    ComboBox,
    BodyLabel,
    PushButton,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.revenue_accrual_worker import RevenueAccrualWorker
from ui.services.app_context import get_app_context
import tools.util as util_module
import core.project_status as ps_module
import core.revenue_accrual as ra_module


# ──────────────────────────────────────────────────────────────
# Inline mini-workers for Project Scope background operations
# ──────────────────────────────────────────────────────────────

class _CountWorker(QThread):
    result = Signal(int)
    error  = Signal(str)

    def __init__(self, pkey: str):
        super().__init__()
        self.pkey = pkey

    def run(self):
        try:
            ra_module.HEADERS = util_module.set_headers()
            self.result.emit(int(ra_module.get_search_option_project_total(self.pkey)))
        except Exception as e:
            self.error.emit(str(e))


class _DownloadSearchListWorker(QThread):
    done = Signal(str)  # "success" or "error:<msg>"

    def __init__(self, pkey: str):
        super().__init__()
        self.pkey = pkey

    def run(self):
        try:
            ra_module.HEADERS = util_module.set_headers()
            ra_module.download_serch_options_list(self.pkey)
            self.done.emit("success")
        except Exception as e:
            self.done.emit(f"error:{e}")


# ──────────────────────────────────────────────────────────────
# Style helpers
# ──────────────────────────────────────────────────────────────

def _combo_style() -> str:
    return f"""
        ComboBox {{
            padding: 6px;
            border: 1px solid {ThemeColors.border_primary()};
            border-radius: 4px;
            background-color: {ThemeColors.background_input()};
            color: {ThemeColors.text_primary()};
        }}
        ComboBox::drop-down {{
            border: none;
            padding-right: 8px;
        }}
        ComboBox::down-arrow {{
            image: none;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 4px solid {ThemeColors.text_secondary()};
            margin-right: 4px;
        }}
    """


_SECTION_LABEL_STYLE = "font-weight: 600; font-size: 13px;"
_MUTED_VALUE_STYLE   = f"font-size: 12px; color: {ThemeColors.text_muted()};"


# ──────────────────────────────────────────────────────────────
# Page
# ──────────────────────────────────────────────────────────────

class RevenueAccrualTaskPage(BaseTaskPage):
    """Task page for revenue accrual generation."""

    def __init__(self, parent=None):
        self.period_data_map: Dict[str, str] = {}
        self._count_worker: Optional[_CountWorker] = None
        self._dl_worker: Optional[_DownloadSearchListWorker] = None

        super().__init__("revenue_accrual", "Rev-Gen Step 1", parent)

        get_app_context().login_state_changed.connect(self._on_login_state_changed)

    # ── BaseTaskPage overrides ─────────────────────────────────

    def _get_page_description(self) -> str:
        return (
            "Generate earned revenue accrual workbooks by combining invoices, "
            "general ledger data, budgets, and labor information."
        )

    def _create_config_widgets(self) -> None:
        self._build_sop()
        self.config_layout.addSpacing(16)
        self._build_period()
        self.config_layout.addSpacing(16)
        self._build_scope()
        self.config_layout.addSpacing(12)

    def _load_config(self) -> None:
        pass

    def _validate_params(self) -> Tuple[bool, str]:
        current_text = self.period_combo.currentText()
        if not current_text or current_text.startswith("Error loading"):
            return False, "Please select an accounting period."
        if current_text not in self.period_data_map:
            return False, "Invalid period selection."
        return True, ""

    def _get_params(self) -> Dict[str, Any]:
        current_text = self.period_combo.currentText()
        period = self.period_data_map.get(current_text, "")
        return {
            "period":      str(period) if period else "",
            "pkey":        self._selected_pkey() or "",
            "option_name": self._selected_option_name() or "",
        }

    def _create_worker(self, params: Dict[str, Any]) -> RevenueAccrualWorker:
        return RevenueAccrualWorker(params)

    # ── Block 1: SOP ──────────────────────────────────────────

    def _build_sop(self) -> None:
        header = BodyLabel("How to Use")
        header.setStyleSheet(_SECTION_LABEL_STYLE)
        self.config_layout.addWidget(header)

        body = BodyLabel(
            "1. Select an accounting period from the dropdown.\n"
            "2. Select an existing Project Scope, or choose '+ Add New' to generate one automatically.\n"
            "3. Click Execute to generate the revenue accrual workbooks."
        )
        body.setWordWrap(True)
        body.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(body)

    # ── Block 2: Accounting Period ─────────────────────────────

    def _build_period(self) -> None:
        header = BodyLabel("Accounting Period")
        header.setStyleSheet(_SECTION_LABEL_STYLE)
        self.config_layout.addWidget(header)

        self.period_combo = ComboBox()
        self.period_combo.setStyleSheet(_combo_style())
        self.period_combo.setMinimumWidth(300)
        self.config_layout.addWidget(self.period_combo)

        self._load_periods()

    # ── Block 3: Project Scope ─────────────────────────────────

    def _build_scope(self) -> None:
        header = BodyLabel("Project Scope")
        header.setStyleSheet(_SECTION_LABEL_STYLE)
        self.config_layout.addWidget(header)

        row = QHBoxLayout()
        row.setSpacing(12)
        row.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.search_option_combo = ComboBox()
        self.search_option_combo.setStyleSheet(_combo_style())
        self.search_option_combo.setMinimumWidth(220)
        self.search_option_combo.currentIndexChanged.connect(self._on_search_option_changed)
        row.addWidget(self.search_option_combo)

        # Right stacked panel: page 0 = existing scope, page 1 = Add New note
        self.search_right_panel = QStackedWidget()

        # Page 0: Total count + Download List button
        existing_page = QWidget()
        ex_layout = QHBoxLayout(existing_page)
        ex_layout.setContentsMargins(0, 0, 0, 0)
        ex_layout.setSpacing(12)
        ex_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.total_projects_label = BodyLabel("Total: —")
        self.total_projects_label.setStyleSheet(_MUTED_VALUE_STYLE)
        ex_layout.addWidget(self.total_projects_label)

        self.dl_list_btn = PushButton("Download List")
        self.dl_list_btn.setMinimumWidth(120)
        self.dl_list_btn.clicked.connect(self._on_download_search_list_clicked)
        ex_layout.addWidget(self.dl_list_btn)

        # Page 1: note that scope is generated automatically on Execute
        add_new_page = QWidget()
        an_layout = QHBoxLayout(add_new_page)
        an_layout.setContentsMargins(0, 0, 0, 0)
        an_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        add_new_note = BodyLabel("A new project scope will be generated automatically when you click Execute.")
        add_new_note.setStyleSheet(f"font-size: 11px; color: {ThemeColors.text_muted()};")
        add_new_note.setWordWrap(True)
        an_layout.addWidget(add_new_note)

        self.search_right_panel.addWidget(existing_page)  # index 0
        self.search_right_panel.addWidget(add_new_page)   # index 1

        row.addWidget(self.search_right_panel)
        row.addStretch()

        self.config_layout.addLayout(row)
        self._load_search_options()

    # ── Period helpers ─────────────────────────────────────────

    def _load_periods(self) -> None:
        if not hasattr(self, "period_combo"):
            return
        try:
            ps_module.HEADERS = util_module.set_headers()
            period_data = ps_module.print_period()

            self.period_combo.clear()
            self.period_data_map.clear()

            for p in period_data:
                display = (
                    f"{p['Period']} | "
                    f"From: {p['AccountPdStart'][:10]} "
                    f"To: {p['AccountPdEnd'][:10]}"
                )
                self.period_data_map[display] = str(p["Period"])
                self.period_combo.addItem(display)

            if self.period_combo.count() > 0:
                self.period_combo.setCurrentIndex(0)

        except Exception as e:
            self.period_combo.clear()
            self.period_combo.addItem("Error loading periods — please check connection")
            print(f"Error loading periods: {e}")

    # ── Search option helpers ──────────────────────────────────

    def _load_search_options(self) -> None:
        if not hasattr(self, "search_option_combo"):
            return
        try:
            ra_module.HEADERS = util_module.set_headers()
            items = ra_module.get_search_list()

            self.search_option_combo.blockSignals(True)
            self.search_option_combo.clear()
            for item in items:
                self.search_option_combo.addItem(item["Name"], userData=item["PKey"])
            self.search_option_combo.addItem("+ Add New", userData=None)
            self.search_option_combo.blockSignals(False)

            if self.search_option_combo.count() > 0:
                self.search_option_combo.blockSignals(True)
                self.search_option_combo.setCurrentIndex(0)
                self.search_option_combo.blockSignals(False)
                self._on_search_option_changed(0)

        except Exception as e:
            print(f"Error loading search options: {e}")
            self.search_option_combo.blockSignals(True)
            self.search_option_combo.clear()
            self.search_option_combo.addItem("+ Add New", userData=None)
            self.search_option_combo.blockSignals(False)
            self.search_right_panel.setCurrentIndex(1)

    def _selected_pkey(self) -> Optional[str]:
        idx = self.search_option_combo.currentIndex()
        return self.search_option_combo.itemData(idx) if idx >= 0 else None

    def _selected_option_name(self) -> Optional[str]:
        text = self.search_option_combo.currentText()
        return None if text == "+ Add New" else text

    # ── Slots ──────────────────────────────────────────────────

    @Slot(bool)
    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        if is_logged_in:
            self._load_periods()
            self._load_search_options()

    @Slot(int)
    def _on_search_option_changed(self, index: int) -> None:
        pkey = self.search_option_combo.itemData(index)
        if pkey is None:
            self.search_right_panel.setCurrentIndex(1)
            return

        self.search_right_panel.setCurrentIndex(0)
        self.total_projects_label.setText("Total: loading...")
        self.dl_list_btn.setEnabled(False)

        if self._count_worker and self._count_worker.isRunning():
            self._count_worker.terminate()
            self._count_worker.wait(500)

        self._count_worker = _CountWorker(pkey)
        self._count_worker.result.connect(self._on_count_loaded)
        self._count_worker.error.connect(self._on_count_error)
        self._count_worker.start()

    @Slot(int)
    def _on_count_loaded(self, count: int) -> None:
        self.total_projects_label.setText(f"Total: {count}")
        self.dl_list_btn.setEnabled(True)

    @Slot(str)
    def _on_count_error(self, msg: str) -> None:
        self.total_projects_label.setText("Total: (error)")
        self.dl_list_btn.setEnabled(False)
        self._append_log("WARNING", f"Could not fetch project count: {msg}")

    @Slot()
    def _on_download_search_list_clicked(self) -> None:
        pkey = self._selected_pkey()
        if not pkey:
            self._append_log("WARNING", "No scope selected.")
            return
        if self._dl_worker and self._dl_worker.isRunning():
            return
        self.dl_list_btn.setEnabled(False)
        self._append_log("INFO", "Downloading project list for selected scope...")
        self._dl_worker = _DownloadSearchListWorker(pkey)
        self._dl_worker.done.connect(self._on_download_search_list_done)
        self._dl_worker.start()

    @Slot(str)
    def _on_download_search_list_done(self, result: str) -> None:
        self.dl_list_btn.setEnabled(True)
        if result == "success":
            from pathlib import Path
            self._append_log("SUCCESS", f"Project list downloaded to {Path.home() / 'Downloads'}")
        else:
            self._append_log("ERROR", f"Download failed: {result.removeprefix('error:')}")

