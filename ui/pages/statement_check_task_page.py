"""
Statement Check task page with modern Fluent Design.

Task: Check statement status for invoices.
"""
from typing import Dict, Any, Tuple, List
from io import BytesIO

from PySide6.QtWidgets import (
    QHBoxLayout,
    QVBoxLayout,
    QFrame,
    QLabel,
    QApplication,
)
from PySide6.QtCore import Qt, QThread, Signal, QBuffer, QIODevice
from PySide6.QtGui import QImage, QPixmap, QKeySequence
from PIL import Image
from qfluentwidgets import (
    ComboBox,
    LineEdit,
    BodyLabel,
    PushButton,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.statement_check_worker import StatementCheckWorker
import core.statement_check as statement_check_module
import tools.util as util_module


def _qimage_to_pil(qimg: QImage) -> Image.Image:
    """Convert a QImage into a PIL Image via an in-memory PNG buffer."""
    buffer = QBuffer()
    buffer.open(QIODevice.OpenModeFlag.ReadWrite)
    qimg.save(buffer, "PNG")
    pil_image = Image.open(BytesIO(bytes(buffer.data()))).copy()
    buffer.close()
    return pil_image


class _OcrWorker(QThread):
    """Run OCR on pasted images off the GUI thread so the UI stays responsive."""

    finished_ocr = Signal(list)  # list[str] of detected numbers

    def __init__(self, images: List[QImage], parent=None):
        super().__init__(parent)
        self._images = images

    def run(self) -> None:
        numbers: List[str] = []
        for qimg in self._images:
            try:
                pil_image = _qimage_to_pil(qimg)
                numbers.extend(statement_check_module.retrieve_numbers_from_image(pil_image))
            except Exception as exc:  # keep going on a bad image
                print(f"OCR failed for one image: {exc}")
        self.finished_ocr.emit(numbers)


class ImagePasteArea(QFrame):
    """
    A focusable drop/paste target that collects invoice images.

    Users click the area then press Ctrl+V to paste a screenshot, or drag &
    drop image files. Each collected image is shown as a thumbnail.
    """

    images_changed = Signal(int)  # current image count

    def __init__(self, parent=None):
        super().__init__(parent)
        self._images: List[QImage] = []

        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumHeight(96)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        self._hint = BodyLabel(
            "Click here, then press Ctrl+V to paste invoice image(s) — "
            "or drag & drop image files."
        )
        self._hint.setWordWrap(True)
        self._hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._hint)

        self._thumb_row = QHBoxLayout()
        self._thumb_row.setSpacing(6)
        self._thumb_row.addStretch()
        layout.addLayout(self._thumb_row)

        self._apply_style()

    def _apply_style(self) -> None:
        border = ThemeColors.border_primary()
        bg = ThemeColors.background_tutorial()
        text = ThemeColors.text_secondary()
        self.setStyleSheet(f"""
            ImagePasteArea {{
                border: 1px dashed {border};
                border-radius: 6px;
                background-color: {bg};
            }}
            ImagePasteArea:focus {{
                border: 1px dashed #4caf50;
            }}
            BodyLabel {{ color: {text}; }}
        """)

    # --- public API -----------------------------------------------------
    def get_images(self) -> List[QImage]:
        return list(self._images)

    def clear_images(self) -> None:
        self._images.clear()
        self._rebuild_thumbnails()
        self.images_changed.emit(0)

    # --- collecting images ----------------------------------------------
    def _add_qimage(self, qimg: QImage) -> bool:
        if qimg is None or qimg.isNull():
            return False
        self._images.append(qimg)
        return True

    def _add_from_mime(self, mime) -> int:
        """Add images from clipboard/drop mime data. Returns count added."""
        added = 0
        if mime.hasImage():
            qimg = mime.imageData()
            if isinstance(qimg, QImage) and self._add_qimage(qimg):
                added += 1
        if mime.hasUrls():
            for url in mime.urls():
                path = url.toLocalFile()
                if path:
                    qimg = QImage(path)
                    if self._add_qimage(qimg):
                        added += 1
        if added:
            self._rebuild_thumbnails()
            self.images_changed.emit(len(self._images))
        return added

    def _rebuild_thumbnails(self) -> None:
        # Remove existing thumbnail widgets (keep the trailing stretch)
        while self._thumb_row.count() > 1:
            item = self._thumb_row.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for qimg in self._images:
            thumb = QLabel()
            pixmap = QPixmap.fromImage(qimg).scaled(
                56, 56,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            thumb.setPixmap(pixmap)
            thumb.setFixedSize(56, 56)
            thumb.setStyleSheet(
                f"border: 1px solid {ThemeColors.border_primary()}; border-radius: 4px;"
            )
            self._thumb_row.insertWidget(self._thumb_row.count() - 1, thumb)

        count = len(self._images)
        if count:
            self._hint.setText(
                f"{count} image(s) ready. Paste more with Ctrl+V, "
                "or click 'Retrieve Numbers from Image'."
            )
        else:
            self._hint.setText(
                "Click here, then press Ctrl+V to paste invoice image(s) — "
                "or drag & drop image files."
            )

    # --- events ---------------------------------------------------------
    def mousePressEvent(self, event) -> None:
        self.setFocus()
        super().mousePressEvent(event)

    def keyPressEvent(self, event) -> None:
        if event.matches(QKeySequence.StandardKey.Paste):
            self._add_from_mime(QApplication.clipboard().mimeData())
            event.accept()
            return
        super().keyPressEvent(event)

    def dragEnterEvent(self, event) -> None:
        mime = event.mimeData()
        if mime.hasImage() or mime.hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        if self._add_from_mime(event.mimeData()):
            event.acceptProposedAction()


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
            "3. Enter invoice numbers separated by commas (e.g., INV001, INV002),\n"
            "   or paste invoice image(s) below and click 'Retrieve Numbers from Image'.\n"
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

        self.config_layout.addSpacing(12)

        # Extract-from-image section
        image_label = BodyLabel("Extract Numbers from Image (OCR)")
        image_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(image_label)

        self.image_paste_area = ImagePasteArea()
        self.config_layout.addWidget(self.image_paste_area)

        image_actions = QHBoxLayout()
        image_actions.setSpacing(8)

        self.retrieve_btn = PushButton("Retrieve Numbers from Image")
        self.retrieve_btn.clicked.connect(self._on_retrieve_from_image)
        image_actions.addWidget(self.retrieve_btn)

        self.clear_images_btn = PushButton("Clear Images")
        self.clear_images_btn.clicked.connect(self.image_paste_area.clear_images)
        image_actions.addWidget(self.clear_images_btn)

        image_actions.addStretch()
        self.config_layout.addLayout(image_actions)

        # Track the OCR helper thread so it isn't garbage-collected mid-run
        self._ocr_worker: _OcrWorker | None = None

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

    def _on_retrieve_from_image(self) -> None:
        """Run OCR on the pasted images and populate the invoice input box."""
        images = self.image_paste_area.get_images()
        if not images:
            self._append_log("WARNING", "No images to process. Paste invoice image(s) first.")
            return

        # Guard against overlapping runs
        if self._ocr_worker is not None and self._ocr_worker.isRunning():
            self._append_log("WARNING", "Still processing the previous image(s). Please wait.")
            return

        self.retrieve_btn.setEnabled(False)
        self.retrieve_btn.setText("Processing...")
        self._append_log("INFO", f"Extracting numbers from {len(images)} image(s)...")

        self._ocr_worker = _OcrWorker(images, parent=self)
        self._ocr_worker.finished_ocr.connect(self._on_ocr_finished)
        self._ocr_worker.start()

    def _on_ocr_finished(self, numbers: List[str]) -> None:
        """Merge OCR-detected numbers into the invoice input box."""
        self.retrieve_btn.setEnabled(True)
        self.retrieve_btn.setText("Retrieve Numbers from Image")

        if not numbers:
            self._append_log("WARNING", "No numbers detected in the pasted image(s).")
            return

        # Merge with any existing numbers, de-duplicating while preserving order
        existing = [inv.strip() for inv in self.invoice_edit.text().split(",") if inv.strip()]
        merged = list(existing)
        for num in numbers:
            if num not in merged:
                merged.append(num)

        self.invoice_edit.setText(", ".join(merged))
        added = len(merged) - len(existing)
        self._append_log(
            "SUCCESS",
            f"Detected {len(numbers)} number(s); added {added} new to the invoice list.",
        )

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
