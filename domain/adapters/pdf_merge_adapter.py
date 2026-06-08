"""
Adapter for the PDF Merge task.

Thin wrapper around core.pdf_merge that routes all log/progress
output through the standard BaseAdapter callback interface so the
UI log panel receives real-time updates.
"""
from typing import Dict, Any, List

from domain.adapters.base_adapter import BaseAdapter
import core.pdf_merge as pdf_merge_module


class PdfMergeAdapter(BaseAdapter):
    """Runs pdf_merge.main() and pipes status back to the UI via callbacks."""

    def execute(self, pdf_paths: List[str] | None = None, **kwargs) -> Dict[str, Any]:  # type: ignore[override]
        """Merge the provided PDF files.

        Args:
            pdf_paths: Ordered list of absolute PDF path strings.

        Returns:
            ``{"success": bool, "message": str}`` — and ``"output_path": str``
            when the merge succeeds.
        """
        paths = pdf_paths or []

        if not paths:
            msg = "No PDF files selected."
            self._log("ERROR", msg)
            return {"success": False, "message": msg}

        self._log("INFO", f"Merging {len(paths)} file(s)…")
        for idx, p in enumerate(paths, 1):
            self._log("INFO", f"  [{idx}/{len(paths)}] {p}")

        self._progress(0, len(paths), "Starting merge…")

        result = pdf_merge_module.main(paths)

        if result.get("success"):
            self._log("SUCCESS", result["message"])
            self._progress(len(paths), len(paths), "Done")
        else:
            self._log("ERROR", result.get("message", "Unknown error during merge."))

        return result
