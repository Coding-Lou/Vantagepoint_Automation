"""
Rev-Gen Report Export task page with Fluent Design.

Layout (top → bottom inside config card):
  1. SOP — compact usage instructions
  2. Accounting Period — Start Period (auto label) + End Period (dropdown)
  3. Project Scope — dropdown | Total count | Download List / Generate
  4. Reports — 2-column checkbox grid
"""
from typing import Dict, Any, Tuple, Optional
from pathlib import Path
import os

from PySide6.QtWidgets import (
    QHBoxLayout,
    QVBoxLayout,
    QGridLayout,
    QCheckBox,
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
    MessageBox,
    MessageDialog,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from ui.utils.log_redirector import LogRedirector
from ui.services.app_context import get_app_context
from workers.report_export_worker import ReportExportWorker
import tools.util as util_module
import core.project_status as ps_module
import core.report_export as re_module

# Output folder — matches core/report_export.py main()
_BASE_FOLDER = Path("C:/temp/revenue_model_export")

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

    def _showComboMenu(self) -> None:
        query = self.text().strip().lower()
        if query:
            filtered = [item for item in self._master_items if query in item.text.lower()]
            original_items = self.items
            original_index = self._currentIndex
            self.items = filtered if filtered else list(self._master_items)
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
# Inline mini-workers for Block 3 background operations
# ──────────────────────────────────────────────────────────────

class _ChangePeriodWorker(QThread):
    done  = Signal(str)   # "ok" or error message

    def __init__(self, period: str):
        super().__init__()
        self.period = period

    def run(self):
        try:
            util_module.change_period(self.period)
            self.done.emit("ok")
        except Exception as e:
            self.done.emit(f"error:{e}")


class _CountWorker(QThread):
    result = Signal(int)
    error  = Signal(str)

    def __init__(self, pkey: str):
        super().__init__()
        self.pkey = pkey

    def run(self):
        try:
            re_module.HEADERS = util_module.set_headers()
            self.result.emit(int(re_module.get_search_option_project_total(self.pkey)))
        except Exception as e:
            self.error.emit(str(e))


class _DownloadSearchListWorker(QThread):
    done = Signal(str)   # "success" or "error:<msg>"

    def __init__(self, base_folder: str, pkey: str):
        super().__init__()
        self.base_folder = base_folder
        self.pkey = pkey

    def run(self):
        try:
            re_module.HEADERS = util_module.set_headers()
            _BASE_FOLDER.mkdir(parents=True, exist_ok=True)
            re_module.download_serch_options_list(self.pkey)
            self.done.emit("success")
        except Exception as e:
            self.done.emit(f"error:{e}")


class _GenerateSearchOptionWorker(QThread):
    done  = Signal(str, str)   # pkey, option_name
    error = Signal(str)
    log   = Signal(str, str)   # level, message  (stdout → Execution Log)

    def __init__(self, end_period: str):
        super().__init__()
        self.end_period = end_period

    def run(self):
        import sys
        old_stdout, old_stderr = sys.stdout, sys.stderr
        redirector = LogRedirector(self.log.emit)
        sys.stdout = redirector
        sys.stderr = redirector
        try:
            util_module.change_period(self.end_period)
            self.log.emit("INFO", f"Active period set to {self.end_period}.")
            re_module.HEADERS = util_module.set_headers()
            re_module.project_list = set()
            # create_new_search_option calls download_GL without base_folder;
            # patch it to supply a default so the call doesn't crash.
            _orig_gl = re_module.download_GL
            def _patched_gl(base_folder="", *args, **kwargs):
                return _orig_gl(base_folder, *args, **kwargs)
            re_module.download_GL = _patched_gl
            try:
                pkey, option_name = re_module.create_new_search_option(self.base_folder, self.end_period)
            finally:
                re_module.download_GL = _orig_gl
            self.done.emit(pkey, option_name)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            redirector.flush()
            sys.stdout = old_stdout
            sys.stderr = old_stderr


# ──────────────────────────────────────────────────────────────
# Shared style helpers
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


_CHECKBOX_STYLE = """
    QCheckBox {
        padding: 4px;
        font-size: 12px;
    }
    QCheckBox::indicator {
        width: 16px;
        height: 16px;
    }
"""

_SECTION_LABEL_STYLE = "font-weight: 600; font-size: 13px;"


# ──────────────────────────────────────────────────────────────
# Page
# ──────────────────────────────────────────────────────────────

class ReportExportTaskPage(BaseTaskPage):
    """Task page for Rev-Gen Report Export."""

    def __init__(self, parent=None):
        self.period_data_map: Dict[str, str] = {}
        self._change_period_worker: Optional[_ChangePeriodWorker] = None
        self._count_worker: Optional[_CountWorker] = None
        self._dl_worker:    Optional[_DownloadSearchListWorker] = None
        self._gen_worker:   Optional[_GenerateSearchOptionWorker] = None

        super().__init__("report_export", "Revenue Model Report Export", parent)

        get_app_context().login_state_changed.connect(self._on_login_state_changed)

    # ── BaseTaskPage overrides ─────────────────────────────────

    def _get_page_description(self) -> str:
        return "Export Revenue Model reports from Vantagepoint ERP for revenue accrual processing."

    def _create_config_widgets(self) -> None:
        self._build_sop()
        self.config_layout.addSpacing(16)
        self._build_periods()
        self.config_layout.addSpacing(16)
        self._build_scope()
        self.config_layout.addSpacing(16)
        self._build_reports()

    def _load_config(self) -> None:
        pass

    def _validate_params(self) -> Tuple[bool, str]:
        if not self._end_period_value():
            return False, "Please select an accounting period."
        if not any(cb.isChecked() for cb in self._checkboxes()):
            return False, "Please select at least one report to download."
        needs_pkey = (
            self.cb_contract.isChecked()
            or self.cb_cost_from_system.isChecked()
            or self.cb_billing.isChecked()
            or self.cb_billed.isChecked()
        )
        if needs_pkey and not self._selected_pkey():
            return (
                False,
                "Contract, Cost From System, Billing, and Billed require a Project Scope. "
                "Please select one or uncheck those reports.",
            )
        return True, ""

    def _get_params(self) -> Dict[str, Any]:
        return {
            "end_period":            self._end_period_value() or "",
            "pkey":                  self._selected_pkey() or "",
            "option_name":           self._selected_option_name() or "",
            "download_project_list": self.cb_project_list.isChecked(),
            "download_estimate":     self.cb_estimate.isChecked(),
            "download_contract":     self.cb_contract.isChecked(),
            "download_cost_gl":      self.cb_cost_gl.isChecked(),
            "download_cost_from_system": self.cb_cost_from_system.isChecked(),
            "download_billing":      self.cb_billing.isChecked(),
            "download_billed":       self.cb_billed.isChecked(),
        }

    def _create_worker(self, params: Dict[str, Any]) -> ReportExportWorker:
        return ReportExportWorker(params)

    # ── Block 1: SOP ──────────────────────────────────────────

    def _build_sop(self) -> None:
        body = BodyLabel(
            "1. Select Start and End Periods using the year/month selectors.\n"
            "2. Select a Project Scope, or click + to create one. Leave empty to auto-generate on Execute.\n"
            "3. Check the reports to export, then click Execute."
        )
        body.setWordWrap(True)
        body.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(body)

    # ── Block 2: Accounting Period ─────────────────────────────

    def _build_periods(self) -> None:
        header = BodyLabel("Accounting Period")
        header.setStyleSheet(_SECTION_LABEL_STYLE)
        self.config_layout.addWidget(header)

        self.period_combo = ComboBox()
        self.period_combo.setStyleSheet(_combo_style())
        self.period_combo.setMinimumWidth(300)
        self.period_combo.currentIndexChanged.connect(self._on_end_period_changed)
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
        self.total_projects_label.setStyleSheet(f"font-size: 12px; color: {ThemeColors.text_muted()};")
        row.addWidget(self.total_projects_label)

        self.dl_list_btn = PushButton("Download List")
        self.dl_list_btn.setMinimumWidth(120)
        self.dl_list_btn.setEnabled(False)
        self.dl_list_btn.clicked.connect(self._on_download_search_list_clicked)
        row.addWidget(self.dl_list_btn)

        row.addStretch()

        self.config_layout.addLayout(row)
        self._load_search_options()

    # ── Block 4: Reports ──────────────────────────────────────

    def _build_reports(self) -> None:
        header = BodyLabel("Reports For Revenue Model")
        header.setStyleSheet(_SECTION_LABEL_STYLE)
        self.config_layout.addWidget(header)

        self.cb_project_list     = QCheckBox("Project List JTD")
        self.cb_estimate         = QCheckBox("Estimate JTD")
        self.cb_contract         = QCheckBox("Contract JTD")
        self.cb_billed           = QCheckBox("Billed JTD")
        self.cb_cost_gl          = QCheckBox("Cost GL JTD")
        self.cb_cost_from_system = QCheckBox("Cost From System JTD")
        self.cb_billing          = QCheckBox("Billing / Spent / WO / Prop JTD")

        for cb in self._checkboxes():
            cb.setChecked(False)
            cb.setStyleSheet(_CHECKBOX_STYLE)

        # 2-column grid matching the mockup order:
        # Col 0 (left)       Col 1 (right)
        # Project List       Cost GL
        # Estimate           Cost From System
        # Contract           Billing / Spent / WO / Prop
        # Billed
        grid = QGridLayout()
        grid.setHorizontalSpacing(32)
        grid.setVerticalSpacing(6)

        grid.addWidget(self.cb_project_list,     0, 0)
        grid.addWidget(self.cb_cost_gl,          0, 1)
        grid.addWidget(self.cb_estimate,         1, 0)
        grid.addWidget(self.cb_cost_from_system, 1, 1)
        grid.addWidget(self.cb_contract,         2, 0)
        grid.addWidget(self.cb_billing,          2, 1)
        grid.addWidget(self.cb_billed,           3, 0)

        self.config_layout.addLayout(grid)

    # ── Period helpers ─────────────────────────────────────────

    def _load_periods(self) -> None:
        if not hasattr(self, "period_combo"):
            return
        try:
            ps_module.HEADERS = util_module.set_headers()
            period_data = ps_module.print_period()

            self.period_combo.blockSignals(True)
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

            self.period_combo.blockSignals(False)

            if self.period_combo.count() > 0:
                self.period_combo.setCurrentIndex(0)
                self._on_end_period_changed(0)

        except Exception as e:
            self.period_combo.blockSignals(False)
            self.period_combo.clear()
            self.period_combo.addItem("Error loading periods — please check connection")
            print(f"Error loading periods: {e}")

    def _end_period_value(self) -> Optional[str]:
        text = self.period_combo.currentText()
        if not text or text.startswith("Error loading"):
            return None
        return self.period_data_map.get(text)

    # ── Search option helpers ──────────────────────────────────

    def _load_search_options(self) -> None:
        if not hasattr(self, "search_option_combo"):
            return
        try:
            re_module.HEADERS = util_module.set_headers()
            items = re_module.get_search_list()

            self.search_option_combo.blockSignals(True)
            self.search_option_combo.clear()
            self.search_option_combo.setText("")
            for item in items:
                self.search_option_combo.addItem(item["Name"], userData=item["PKey"])
            self.search_option_combo.blockSignals(False)

            self.search_option_combo._master_items = list(self.search_option_combo.items)
            self.total_projects_label.setText("Total: —")
            self.dl_list_btn.setEnabled(False)
            self.scope_delete_btn.setEnabled(False)

        except Exception as e:
            print(f"Error loading search options: {e}")
            self.search_option_combo.blockSignals(True)
            self.search_option_combo.clear()
            self.search_option_combo.blockSignals(False)
            self.search_option_combo._master_items = []
            self.total_projects_label.setText("Total: —")
            self.dl_list_btn.setEnabled(False)
            self.scope_delete_btn.setEnabled(False)

    def _selected_pkey(self) -> Optional[str]:
        idx = self.search_option_combo.currentIndex()
        return self.search_option_combo.itemData(idx) if idx >= 0 else None

    def _selected_option_name(self) -> Optional[str]:
        text = self.search_option_combo.currentText()
        return text or None

    # ── Execute override ───────────────────────────────────────

    @Slot()
    def _on_execute_clicked(self) -> None:
        # Show dialog if no reports are checked
        if not any(cb.isChecked() for cb in self._checkboxes()):
            dlg = MessageBox(
                "No Reports Selected",
                "Please select at least one report to download before proceeding.",
                self.window(),
            )
            dlg.exec()
            return

        # Auto-generate scope when no scope is selected
        if self._selected_pkey() is None:
            end = self._end_period_value()
            if self._gen_worker and self._gen_worker.isRunning():
                return
            self.execute_btn.setEnabled(False)
            self._append_log("INFO", f"Generating new project scope for period {end}. This may take a few minutes...")
            self._gen_worker = _GenerateSearchOptionWorker(end)
            self._gen_worker.log.connect(self._append_log)
            self._gen_worker.done.connect(self._on_generate_for_execute_done)
            self._gen_worker.error.connect(self._on_generate_for_execute_error)
            self._gen_worker.start()
            return

        super()._on_execute_clicked()

    @Slot(str, str)
    def _on_generate_for_execute_done(self, pkey: str, option_name: str) -> None:
        self._append_log("SUCCESS", f"Scope created: {option_name}")

        # Add the generated option and select it — no network reload needed.
        combo = self.search_option_combo
        combo.blockSignals(True)
        combo.addItem(option_name, userData=pkey)
        combo._master_items = list(combo.items)
        new_idx = combo.count() - 1
        combo.setCurrentIndex(new_idx)
        combo.blockSignals(False)
        self._on_search_option_changed(new_idx)

        self.execute_btn.setEnabled(True)
        # Do NOT call super()._on_execute_clicked() — it would:
        #   1. Re-run check_login() (HTTP request) even though the session was just used
        #   2. Inside BaseWorker, run _check_login() a second time, potentially
        #      triggering _auto_login() (browser popup) on a transient failure
        #   3. Clear the Execution Log, discarding the generation output
        # We know the session is valid, so skip straight to starting the worker.
        self._launch_export_worker()

    def _launch_export_worker(self) -> None:
        """Start the export worker directly, bypassing the login re-check."""
        is_valid, error_msg = self._validate_params()
        if not is_valid:
            self._append_log("ERROR", f"Validation failed: {error_msg}")
            return

        # Clean up any stale worker
        if self.current_worker:
            if self.current_worker.isRunning():
                self._append_log("WARNING", "Task is already running.")
                return
            try:
                self.current_worker.log_signal.disconnect()
                self.current_worker.progress_signal.disconnect()
                self.current_worker.finished.disconnect()
            except Exception:
                pass
            self.current_worker = None

        params = self._get_params()
        self.current_worker = self._create_worker(params)
        self.current_worker.log_signal.connect(self._on_log_received)
        self.current_worker.progress_signal.connect(self._on_progress_received)
        self.current_worker.finished.connect(self._on_worker_finished)

        self._set_execution_enabled(False)
        # Keep the generate logs — do not clear
        self._append_log("INFO", f"Starting export: {self.task_name}")
        self.task_started.emit(self.task_id)
        self.current_worker.start()

    @Slot(str)
    def _on_generate_for_execute_error(self, msg: str) -> None:
        self._append_log("ERROR", f"Failed to generate project scope: {msg}")
        self.execute_btn.setEnabled(True)

    # ── Slots ──────────────────────────────────────────────────

    @Slot(int)
    def _on_end_period_changed(self, _: int) -> None:
        end = self._end_period_value()
        if not end:
            return
        if self._change_period_worker and self._change_period_worker.isRunning():
            self._change_period_worker.terminate()
            self._change_period_worker.wait(300)
        self._change_period_worker = _ChangePeriodWorker(end)
        self._change_period_worker.done.connect(self._on_change_period_done)
        self._change_period_worker.start()

    @Slot(str)
    def _on_change_period_done(self, result: str) -> None:
        if result != "ok":
            self._append_log("WARNING", f"Could not set active period: {result.removeprefix('error:')}")

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
    def _on_download_search_list_clicked(self) -> None:
        pkey = self._selected_pkey()
        if not pkey:
            self._append_log("WARNING", "No scope selected.")
            return
        if self._dl_worker and self._dl_worker.isRunning():
            return
        self.dl_list_btn.setEnabled(False)
        self._append_log("INFO", "Downloading project list for selected scope...")
        self._dl_worker = _DownloadSearchListWorker(str(_BASE_FOLDER), pkey)
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

    @Slot(bool)
    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        if is_logged_in:
            self._load_periods()
            self._load_search_options()

    @Slot()
    def _on_add_scope_clicked(self) -> None:
        dialog = _NewScopeDialog(self.window())
        if dialog.exec():
            self._append_log("INFO", "New scope saved — refreshing list...")
            self._load_search_options()

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

    # ── Utilities ──────────────────────────────────────────────

    def _checkboxes(self):
        attrs = [
            "cb_project_list", "cb_estimate", "cb_contract", "cb_billed",
            "cb_cost_gl", "cb_cost_from_system", "cb_billing",
        ]
        return [getattr(self, a) for a in attrs if hasattr(self, a)]
