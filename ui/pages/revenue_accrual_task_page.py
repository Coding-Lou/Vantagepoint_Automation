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
    QDialog,
)
from PySide6.QtCore import Qt, Signal, Slot, QThread

from qfluentwidgets import (
    ComboBox,
    EditableComboBox,
    LineEdit,
    CheckBox,
    BodyLabel,
    StrongBodyLabel,
    PushButton,
    PrimaryPushButton,
    ToolButton,
    FluentIcon,
    MessageDialog,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.revenue_accrual_worker import RevenueAccrualWorker
from ui.services.app_context import get_app_context
import tools.util as util_module
import core.project_status as ps_module
import core.revenue_accrual as ra_module


# ──────────────────────────────────────────────────────────────
# Searchable combo box for Project Scope
# ──────────────────────────────────────────────────────────────

class _SearchableScopeComboBox(EditableComboBox):
    """EditableComboBox that filters the dropdown by typed text.

    Filtering is applied at menu-open time by temporarily replacing self.items
    with a filtered subset.  The click handler uses findText() (text-based
    lookup), so restoring self.items before the user clicks is safe.
    Also suppresses the default behaviour of creating a new item when the user
    presses Return on unmatched text.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._master_items: list = []

    def addItem(self, text, icon=None, userData=None) -> None:
        from qfluentwidgets.components.widgets.combo_box import ComboItem
        self.items.append(ComboItem(text, icon, userData))
        # Skip the parent's auto-setCurrentIndex(0) so the field stays empty

    def _showComboMenu(self) -> None:
        query = self.text().strip().lower()
        if query:
            filtered = [item for item in self._master_items if query in item.text.lower()]
            original_items = self.items
            original_index = self._currentIndex
            self.items = filtered if filtered else list(self._master_items)
            # Reset index so ComboBoxBase doesn't try menu.actions()[stale_index]
            # on the smaller filtered list, which raises IndexError and prevents restore.
            self._currentIndex = -1
            try:
                super()._showComboMenu()
            finally:
                self.items = original_items
                self._currentIndex = original_index
        else:
            super()._showComboMenu()

    def _onReturnPressed(self) -> None:
        text = self.text()
        if not text:
            return
        index = self.findText(text)
        if index >= 0 and index != self.currentIndex():
            self._currentIndex = index
            self.currentIndexChanged.emit(index)


# ──────────────────────────────────────────────────────────────
# New Scope dialog
# ──────────────────────────────────────────────────────────────

class _NewScopeDialog(QDialog):
    """Modal form for creating a new Project Scope search option."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Project Scope")
        self.setFixedWidth(440)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(10)

        title = StrongBodyLabel("New Project Scope")
        title.setStyleSheet("font-size: 15px;")
        layout.addWidget(title)
        layout.addSpacing(4)

        layout.addWidget(BodyLabel("Scope Name"))
        self.name_edit = LineEdit()
        self.name_edit.setPlaceholderText("e.g. Q2 2026 Active Projects")
        layout.addWidget(self.name_edit)

        layout.addWidget(BodyLabel("Project List (comma-separated)"))
        self.project_edit = LineEdit()
        self.project_edit.setPlaceholderText("e.g. P001, P002, P003")
        layout.addWidget(self.project_edit)

        self.public_check = CheckBox("Public (visible to all accounting roles)")
        self.public_check.setChecked(True)
        layout.addWidget(self.public_check)

        self.error_label = BodyLabel("")
        self.error_label.setStyleSheet(f"color: {ThemeColors.status_error()}; font-size: 11px;")
        self.error_label.setVisible(False)
        layout.addWidget(self.error_label)

        layout.addSpacing(6)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = PushButton("Cancel")
        self.confirm_btn = PrimaryPushButton("Confirm")
        btn_row.addWidget(self.cancel_btn)
        btn_row.addWidget(self.confirm_btn)
        layout.addLayout(btn_row)

        self.cancel_btn.clicked.connect(self.reject)
        self.confirm_btn.clicked.connect(self._on_confirm)

    def _on_confirm(self) -> None:
        name = self.name_edit.text().strip()
        if not name:
            self.error_label.setText("Scope name is required.")
            self.error_label.setVisible(True)
            return

        projects = [p.strip() for p in self.project_edit.text().split(",") if p.strip()]
        is_public = self.public_check.isChecked()

        try:
            self.confirm_btn.setEnabled(False)
            headers = util_module.set_headers()
            util_module.save_new_search_options(
                header=headers,
                projects=projects,
                saveName=name,
                isPublic=is_public,
            )
            self.accept()
        except Exception as e:
            self.error_label.setText(f"Save failed: {e}")
            self.error_label.setVisible(True)
            self.confirm_btn.setEnabled(True)


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
            "2. Select an existing Project Scope, or click the + button to create a new one.\n"
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
        row.setSpacing(8)
        row.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.search_option_combo = _SearchableScopeComboBox()
        self.search_option_combo.setPlaceholderText("Type to filter scopes...")
        self.search_option_combo.setMinimumWidth(220)
        self.search_option_combo.currentIndexChanged.connect(self._on_search_option_changed)
        row.addWidget(self.search_option_combo)

        self.scope_refresh_btn = ToolButton(FluentIcon.SYNC)
        self.scope_refresh_btn.setFixedSize(32, 32)
        self.scope_refresh_btn.clicked.connect(self._load_search_options)
        row.addWidget(self.scope_refresh_btn)

        self.scope_add_btn = ToolButton(FluentIcon.ADD)
        self.scope_add_btn.setFixedSize(32, 32)
        self.scope_add_btn.clicked.connect(self._on_add_scope_clicked)
        row.addWidget(self.scope_add_btn)

        self.scope_delete_btn = ToolButton(FluentIcon.REMOVE)
        self.scope_delete_btn.setFixedSize(32, 32)
        self.scope_delete_btn.setEnabled(False)
        self.scope_delete_btn.clicked.connect(self._on_delete_scope_clicked)
        row.addWidget(self.scope_delete_btn)

        row.addSpacing(4)

        self.total_projects_label = BodyLabel("Total: —")
        self.total_projects_label.setStyleSheet(_MUTED_VALUE_STYLE)
        row.addWidget(self.total_projects_label)

        self.dl_list_btn = PushButton("Download List")
        self.dl_list_btn.setMinimumWidth(120)
        self.dl_list_btn.setEnabled(False)
        self.dl_list_btn.clicked.connect(self._on_download_search_list_clicked)
        row.addWidget(self.dl_list_btn)

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
            self.search_option_combo.setText("")
            for item in items:
                self.search_option_combo.addItem(item["Name"], userData=item["PKey"])
            self.search_option_combo.blockSignals(False)

            self.search_option_combo._master_items = list(self.search_option_combo.items)
            self.total_projects_label.setText("Total: —")
            self.dl_list_btn.setEnabled(False)

        except Exception as e:
            print(f"Error loading search options: {e}")
            self.search_option_combo.blockSignals(True)
            self.search_option_combo.clear()
            self.search_option_combo.blockSignals(False)
            self.search_option_combo._master_items = []
            self.total_projects_label.setText("Total: —")
            self.dl_list_btn.setEnabled(False)

    def _selected_pkey(self) -> Optional[str]:
        idx = self.search_option_combo.currentIndex()
        return self.search_option_combo.itemData(idx) if idx >= 0 else None

    def _selected_option_name(self) -> Optional[str]:
        text = self.search_option_combo.currentText()
        return text or None

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
            self.total_projects_label.setText("Total: —")
            self.dl_list_btn.setEnabled(False)
            self.scope_delete_btn.setEnabled(False)
            return

        self.scope_delete_btn.setEnabled(True)
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
    def _on_delete_scope_clicked(self) -> None:
        pkey = self._selected_pkey()
        if not pkey:
            return
        name = self._selected_option_name() or "this scope"
        dialog = MessageDialog(
            "Delete Scope",
            f'Delete "{name}"? This cannot be undone.',
            self.window(),
        )
        if dialog.exec():
            try:
                headers = util_module.set_headers()
                util_module.delete_search_options(header=headers, key=pkey)
                self._append_log("INFO", f'Scope "{name}" deleted — refreshing list...')
                self._load_search_options()
            except Exception as e:
                self._append_log("ERROR", f"Delete failed: {e}")

    @Slot()
    def _on_add_scope_clicked(self) -> None:
        dialog = _NewScopeDialog(self.window())
        if dialog.exec():
            self._append_log("INFO", "New scope saved — refreshing list...")
            self._load_search_options()

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

