"""
Startup progress dialog (modern minimal style).

This dialog is meant to be shown briefly during application startup
to provide immediate UI feedback and reduce perceived load time.

All code and comments are in English by request.
"""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QDialog, QVBoxLayout, QProgressBar, QWidget

from qfluentwidgets import CardWidget, TitleLabel, BodyLabel


class StartupProgressDialog(QDialog):
    """A small frameless dialog with a progress bar and stage text."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Starting...")

        # Frameless, minimal footprint, stays on top during startup.
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setModal(False)

        self.setFixedSize(440, 170)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        self._card = CardWidget(self)
        card_layout = QVBoxLayout(self._card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(8)

        self._title = TitleLabel("QCA Accounting Automation Tool")
        self._title.setStyleSheet("font-weight: 600;")
        card_layout.addWidget(self._title)

        self._subtitle = BodyLabel("Starting...")
        self._subtitle.setWordWrap(True)
        self._subtitle.setStyleSheet("color: #808080;")
        card_layout.addWidget(self._subtitle)

        self._progress = QProgressBar(self)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        self._progress.setFixedHeight(10)
        # Minimal rounded bar style; keeps consistent in both themes.
        self._progress.setStyleSheet(
            """
            QProgressBar {
                border: 1px solid rgba(120, 120, 120, 80);
                border-radius: 5px;
                background: rgba(127, 127, 127, 30);
            }
            QProgressBar::chunk {
                border-radius: 5px;
                background: rgba(0, 120, 215, 200);
            }
            """
        )
        card_layout.addWidget(self._progress)

        root.addWidget(self._card)

        self._center_on_screen()

    def _center_on_screen(self) -> None:
        """Center the dialog on the primary screen."""
        screen = QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)

    def set_progress(self, percent: int, text: str) -> None:
        """Update progress percent (0-100) and stage text."""
        p = max(0, min(100, int(percent)))
        self._progress.setValue(p)
        self._subtitle.setText(text or "Starting...")

