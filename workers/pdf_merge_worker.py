"""
Worker for the PDF Merge task.

Overrides BaseWorker.run() to skip the ERP login check — PDF merge
is a purely local file operation that requires no network access or
authenticated session.
"""
from typing import Dict, Any

from workers.base_worker import BaseWorker
from domain.adapters.pdf_merge_adapter import PdfMergeAdapter


class PdfMergeWorker(BaseWorker):
    """Background QThread that merges PDF files without touching the ERP."""

    def __init__(self, params: Dict[str, Any]) -> None:
        super().__init__("pdf_merge", params, auto_login_enabled=False)
        self.adapter = PdfMergeAdapter(
            log_callback=self.log_signal.emit,
            progress_callback=self.progress_signal.emit,
        )

    # ------------------------------------------------------------------
    # Override run() to skip the ERP login gate entirely.
    # The base class run() calls _check_login() / _auto_login(), which
    # requires a live Vantagepoint session.  PDF merge never needs that.
    # ------------------------------------------------------------------

    def run(self) -> None:
        try:
            self._setup_log_redirect()

            if self.is_cancelled:
                self.finished.emit({"success": False, "message": "Task was cancelled"})
                return

            result = self.execute()
            self.finished.emit(result)

        except Exception as exc:
            self.finished.emit({
                "success": False,
                "message": f"PDF merge task exception: {exc}",
            })
        finally:
            self._restore_log_redirect()

    def execute(self) -> Dict[str, Any]:
        if self.is_cancelled:
            return {"success": False, "message": "Task was cancelled"}
        try:
            return self.adapter.execute(pdf_paths=self.params.get("pdf_paths", []))
        except Exception as exc:
            return {"success": False, "message": f"PDF merge failed: {exc}"}
