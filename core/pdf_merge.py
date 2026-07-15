"""
PDF Merge business logic.

Merges multiple PDF files into a single output file using pypdf.
Output is written to the current working directory as merged.pdf,
with automatic filename collision resolution (merged_1.pdf, merged_2.pdf, ...).
No ERP login required — this is a purely local file operation.
"""
from pathlib import Path
from typing import List, Dict, Any
import platform
import subprocess


def get_output_path(directory: Path) -> Path:
    """Return an unused merged-PDF path inside *directory*.

    Tries merged.pdf first, then merged_1.pdf, merged_2.pdf, … until it finds
    a name that does not already exist on disk.
    """
    candidate = directory / "merged.pdf"
    if not candidate.exists():
        return candidate
    n = 1
    while True:
        candidate = directory / f"merged_{n}.pdf"
        if not candidate.exists():
            return candidate
        n += 1


def merge_pdfs(pdf_paths: List[Path], output_path: Path) -> int:
    """Merge *pdf_paths* (in order) into *output_path*.

    Returns the total number of pages written to the output file.
    Raises on any I/O or parse error so callers can log the message.
    """
    from pypdf import PdfWriter, PdfReader

    writer = PdfWriter()
    page_count = 0
    for path in pdf_paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
            page_count += 1

    with open(output_path, "wb") as fh:
        writer.write(fh)

    return page_count


def open_folder(folder: Path) -> None:
    """Open *folder* in the native OS file manager.

    Silently swallows errors — failure to open a window must never crash the app.
    """
    try:
        system = platform.system()
        if system == "Windows":
            subprocess.Popen(["explorer", str(folder)])
        elif system == "Darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        pass


def main(pdf_paths: List[str] | None = None) -> Dict[str, Any]:
    """Entry point for both CLI and adapter use.

    Args:
        pdf_paths: Ordered list of absolute PDF path strings.

    Returns:
        ``{"success": bool, "message": str, "output_path": str}``
        where ``output_path`` is only present on success.
    """
    if not pdf_paths:
        return {"success": False, "message": "No PDF files provided."}

    paths: List[Path] = []
    for raw in pdf_paths:
        p = Path(raw)
        if not p.exists():
            return {"success": False, "message": f"File not found: {p}"}
        if p.suffix.lower() != ".pdf":
            return {"success": False, "message": f"Not a PDF file: {p.name}"}
        paths.append(p)

    output_path = get_output_path(Path.cwd())
    try:
        page_count = merge_pdfs(paths, output_path)
    except Exception as exc:
        return {"success": False, "message": f"Merge failed: {exc}"}

    return {
        "success": True,
        "message": (
            f"Merged {len(paths)} file(s) → {output_path}"
            f" ({page_count} page{'s' if page_count != 1 else ''} total)"
        ),
        "output_path": str(output_path),
    }
