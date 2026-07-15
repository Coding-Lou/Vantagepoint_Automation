"""
PDF Merge task page with Fluent Design.

Lets users collect PDF files via file picker or drag-and-drop,
reorder them, and merge them into a single output file — all without
requiring an ERP login (purely local operation).

Architecture notes
------------------
* Subclasses BaseTaskPage to inherit the standard layout (config card,
  log card, Execute/Cancel action bar) and all log/signal machinery.
* Overrides _on_execute_clicked() to skip the ERP login gate.
* Overrides _on_worker_finished() to auto-open the output folder after
  the MainWindow completion dialog is dismissed.
* Uses _DropListWidget (a QListWidget subclass) that accepts both
  internal drag-to-reorder and external PDF-file drops from the OS.
  The widget owns its own styling and hint-label logic so the page
  stays thin.
"""
from pathlib import Path
from typing import Dict, Any, List, Tuple

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QFileDialog,
)
from PySide6.QtCore import Qt, Slot, QUrl
from PySide6.QtGui import (
    QDragEnterEvent,
    QDragMoveEvent,
    QDropEvent,
    QDesktopServices,
)

from qfluentwidgets import BodyLabel, PushButton

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.pdf_merge_worker import PdfMergeWorker
import core.pdf_merge as pdf_merge_module


# ---------------------------------------------------------------------------
# Drop-enabled list widget
# ---------------------------------------------------------------------------

class _DropListWidget(QListWidget):
    """QListWidget with dual drag-drop support, an empty-state hint label,
    and drag-over border highlight.

    Features
    --------
    * Internal drag: rows can be dragged to reorder (Qt InternalMove).
    * External drag: PDF files dragged from the OS file manager are appended,
      skipping duplicates.  Only .pdf extensions are accepted.
    * Hint label: shows "Drag & Drop PDF files here / or click Add Files"
      centered inside the viewport whenever the list is empty; disappears
      automatically when the first item is added and reappears when the last
      item is removed.
    * Drop-zone border: dashed border indicates the area is a drag target;
      turns solid blue (#0078d4) while external files are being dragged over.
    """

    # Windows / Fluent accent colour used for the active-drag highlight.
    _DRAG_ACTIVE_BORDER = "#0078d4"

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        # InternalMove lets Qt handle row reordering natively.
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        # Apply the normal (non-drag) stylesheet immediately.
        self._apply_normal_style()

        # ── Hint label ────────────────────────────────────────────────────
        # Parented to the viewport so it overlays the list content area.
        self._hint = QLabel(self.viewport())
        self._hint.setText(
            "Drag & Drop PDF files here\nor click  Add Files"
        )
        self._hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Mouse events pass through so clicks still reach the list.
        self._hint.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._hint.setWordWrap(True)
        self._refresh_hint_style()
        self._hint.raise_()  # Keep the label above any item rows.

        # Connect model row-change signals → update hint visibility.
        self.model().rowsInserted.connect(self._on_count_changed)
        self.model().rowsRemoved.connect(self._on_count_changed)

        # Set initial hint visibility (list is empty at construction).
        self._on_count_changed()

    # ── Qt event overrides ────────────────────────────────────────────────

    def resizeEvent(self, event) -> None:
        """Keep the hint label filling the entire viewport."""
        super().resizeEvent(event)
        if hasattr(self, "_hint"):
            self._hint.setGeometry(self.viewport().rect())

    # -- External (OS file-manager) drop events ---------------------------

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            self._apply_drag_active_style()
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            super().dragMoveEvent(event)

    def dragLeaveEvent(self, event) -> None:
        """Restore normal border when the drag leaves without dropping."""
        self._apply_normal_style()
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self._apply_normal_style()
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(".pdf"):
                    self._add_unique(file_path)
            event.accept()
        else:
            # Internal reorder — delegate to Qt's built-in handler.
            super().dropEvent(event)

    # ── Public helpers ────────────────────────────────────────────────────

    def add_path(self, path: str) -> bool:
        """Append *path* to the list; return True if it was not already present."""
        return self._add_unique(path)

    def ordered_paths(self) -> List[str]:
        """Return every item's stored path in current display order."""
        return [
            self.item(i).data(Qt.UserRole)
            for i in range(self.count())
            if self.item(i).data(Qt.UserRole)
        ]

    # ── Private ───────────────────────────────────────────────────────────

    def _add_unique(self, path: str) -> bool:
        """Append *path* as a new list row, silently ignoring duplicates."""
        for i in range(self.count()):
            if self.item(i).data(Qt.UserRole) == path:
                return False
        item = QListWidgetItem(Path(path).name)
        item.setData(Qt.UserRole, path)
        item.setToolTip(path)
        self.addItem(item)
        return True

    @Slot()
    def _on_count_changed(self) -> None:
        """Show the hint only while the list contains no items."""
        self._hint.setVisible(self.count() == 0)

    # ── Styling ───────────────────────────────────────────────────────────

    def _apply_normal_style(self) -> None:
        """Apply the resting drop-zone stylesheet (dashed border)."""
        self.setStyleSheet(self._build_style(active_drag=False))

    def _apply_drag_active_style(self) -> None:
        """Apply the drag-over stylesheet (solid accent-blue border)."""
        self.setStyleSheet(self._build_style(active_drag=True))

    def _build_style(self, *, active_drag: bool) -> str:
        """Construct a complete stylesheet for either resting or drag-active state.

        The border is always dashed during rest to communicate that this area
        is a drop target; it changes to a solid accent-colour stroke when the
        user is actively dragging files over it.
        """
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border_colour = ThemeColors.border_primary()
        hover = ThemeColors.background_menu_hover()

        if active_drag:
            border_css = f"2px solid {self._DRAG_ACTIVE_BORDER}"
            radius = "6px"
        else:
            border_css = f"1px dashed {border_colour}"
            radius = "6px"

        return f"""
            QListWidget {{
                background-color: {bg};
                color: {text};
                border: {border_css};
                border-radius: {radius};
                font-family: 'Segoe UI', sans-serif;
                font-size: 13px;
                padding: 4px;
                outline: none;
            }}
            QListWidget::item {{
                padding: 7px 10px;
                border-radius: 3px;
            }}
            QListWidget::item:selected {{
                background-color: {self._DRAG_ACTIVE_BORDER};
                color: #ffffff;
            }}
            QListWidget::item:hover:!selected {{
                background-color: {hover};
            }}
            QScrollBar:vertical {{
                width: 6px;
                background: transparent;
            }}
            QScrollBar::handle:vertical {{
                background: {border_colour};
                border-radius: 3px;
            }}
        """

    def _refresh_hint_style(self) -> None:
        """Apply a theme-aware style to the overlay hint label."""
        self._hint.setStyleSheet(f"""
            QLabel {{
                color: {ThemeColors.text_muted()};
                font-size: 13px;
                font-family: 'Segoe UI', sans-serif;
                background: transparent;
                line-height: 1.8;
                padding: 20px;
            }}
        """)


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

class PdfMergePage(BaseTaskPage):
    """Task page for merging multiple PDFs into one file.

    Registered in MainWindow under the task_id ``"pdf_merge"`` so it
    participates in the standard page-disable / re-enable lifecycle when
    any background task runs.
    """

    def __init__(self, parent=None) -> None:
        super().__init__("pdf_merge", "PDF Merge", parent)
        # Rename the Execute button now that the base class has created it.
        self.execute_btn.setText("Merge PDFs")

    # -- BaseTaskPage hooks -----------------------------------------------

    def _get_page_description(self) -> str:
        return (
            "Combine multiple PDF files into a single document. "
            "Drag-and-drop files into the list, or use Add Files to browse."
        )

    def _create_config_widgets(self) -> None:
        """Build the file-list section inside the config card."""

        # Tutorial box
        tutorial = BodyLabel(
            "1. Add PDF files with 'Add Files' or drag them from the file manager.\n"
            "2. Drag rows within the list, or use Move Up / Move Down to reorder.\n"
            "3. Double-click a row to preview the PDF in your default viewer.\n"
            "4. Click 'Merge PDFs' — the result is saved to the current working directory."
        )
        tutorial.setWordWrap(True)
        tutorial.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial)

        self.config_layout.addSpacing(12)

        # Section label
        files_label = BodyLabel("PDF Files to Merge")
        files_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(files_label)

        # Drop list — styling and hint label are managed entirely by the widget.
        self.file_list = _DropListWidget()
        self.file_list.setMinimumHeight(240)
        self.file_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.config_layout.addWidget(self.file_list)

        # Control button row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.add_btn = PushButton("Add Files")
        self.add_btn.clicked.connect(self._on_add_files)
        btn_row.addWidget(self.add_btn)

        self.remove_btn = PushButton("Remove")
        self.remove_btn.clicked.connect(self._on_remove_selected)
        btn_row.addWidget(self.remove_btn)

        self.clear_btn = PushButton("Clear All")
        self.clear_btn.clicked.connect(self._on_clear_all)
        btn_row.addWidget(self.clear_btn)

        # Move buttons pushed to the right
        btn_row.addStretch()

        self.move_up_btn = PushButton("Move Up")
        self.move_up_btn.clicked.connect(self._on_move_up)
        btn_row.addWidget(self.move_up_btn)

        self.move_down_btn = PushButton("Move Down")
        self.move_down_btn.clicked.connect(self._on_move_down)
        btn_row.addWidget(self.move_down_btn)

        self.config_layout.addLayout(btn_row)

    def _load_config(self) -> None:
        """No persistent config for PDF merge — nothing to load."""
        pass

    def _validate_params(self) -> Tuple[bool, str]:
        count = self.file_list.count()
        if count == 0:
            return False, "Please add at least one PDF file."
        if count < 2:
            return False, "Please add at least two PDF files to merge."
        return True, ""

    def _get_params(self) -> Dict[str, Any]:
        return {"pdf_paths": self.file_list.ordered_paths()}

    def _create_worker(self, params: Dict[str, Any]) -> PdfMergeWorker:
        return PdfMergeWorker(params)

    # -- Execute override (skip ERP login check) --------------------------

    @Slot()
    def _on_execute_clicked(self) -> None:
        """Start the merge worker, bypassing the ERP login gate.

        PDF merge is a local-only operation — no Vantagepoint session is
        required.  This replaces BaseTaskPage._on_execute_clicked() entirely,
        preserving all other worker lifecycle logic unchanged.
        """
        is_valid, error_msg = self._validate_params()
        if not is_valid:
            self._append_log("ERROR", f"Parameter validation failed: {error_msg}")
            return

        # Clean up any previously completed (but not yet garbage-collected) worker.
        if self.current_worker:
            if self.current_worker.isRunning():
                self._append_log("WARNING", "Task is already running. Please wait for completion.")
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
        self.log_viewer.clear()
        self._append_log("INFO", f"Starting task: {self.task_name}")

        # Notify MainWindow so it disables all other task pages.
        self.task_started.emit(self.task_id)

        self.current_worker.start()

    # -- Post-merge: open output folder automatically --------------------

    @Slot(dict)
    def _on_worker_finished(self, result: Dict[str, Any]) -> None:
        """Delegate to base (which shows the completion dialog), then open folder.

        The base class emits task_finished → MainWindow shows a modal MessageBox.
        After the user dismisses that dialog, execution returns here and we open
        the output folder — so the folder appears after the dialog, not during it.
        """
        super()._on_worker_finished(result)

        if result.get("success"):
            output_path = result.get("output_path", "")
            if output_path:
                try:
                    pdf_merge_module.open_folder(Path(output_path).parent)
                except Exception:
                    pass

    # -- Slot handlers ----------------------------------------------------

    @Slot()
    def _on_add_files(self) -> None:
        """Open a multi-file PDF picker and append selected files."""
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF Files",
            "",
            "PDF Files (*.pdf);;All Files (*)",
        )
        for path in paths:
            self.file_list.add_path(path)

    @Slot()
    def _on_remove_selected(self) -> None:
        """Remove the currently selected row from the file list."""
        row = self.file_list.currentRow()
        if row >= 0:
            self.file_list.takeItem(row)
            # Keep a sensible selection after removal.
            new_count = self.file_list.count()
            if new_count > 0:
                self.file_list.setCurrentRow(min(row, new_count - 1))

    @Slot()
    def _on_clear_all(self) -> None:
        """Remove all items from the file list."""
        self.file_list.clear()

    @Slot()
    def _on_move_up(self) -> None:
        """Shift the selected row one position towards the top."""
        row = self.file_list.currentRow()
        if row > 0:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row - 1, item)
            self.file_list.setCurrentRow(row - 1)

    @Slot()
    def _on_move_down(self) -> None:
        """Shift the selected row one position towards the bottom."""
        row = self.file_list.currentRow()
        if 0 <= row < self.file_list.count() - 1:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row + 1, item)
            self.file_list.setCurrentRow(row + 1)

    @Slot(QListWidgetItem)
    def _on_item_double_clicked(self, item: QListWidgetItem) -> None:
        """Open the PDF at *item*'s path in the system default viewer."""
        path = item.data(Qt.UserRole)
        if path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
