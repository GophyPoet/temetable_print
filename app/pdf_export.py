"""
PDF export module.

Uses WeasyPrint to convert rendered HTML into print-ready PDF.
"""

import os
import tempfile
from weasyprint import HTML

from .parser import SheetData
from .renderer import render_full_document


def _get_base_url():
    """Get base URL for WeasyPrint to resolve relative paths."""
    return os.path.dirname(os.path.abspath(__file__))


def generate_combined_pdf(sheets: list[SheetData]) -> bytes:
    """
    Generate a single PDF containing all printable sheets.

    Each sheet becomes exactly one A4 page.
    Returns PDF as bytes.
    """
    html_content = render_full_document(sheets)
    html_doc = HTML(string=html_content, base_url=_get_base_url())
    return html_doc.write_pdf()


def generate_separate_pdfs(sheets: list[SheetData]) -> list[tuple[str, bytes]]:
    """
    Generate separate PDFs, one per printable sheet.

    Returns list of (filename, pdf_bytes) tuples.
    """
    printable = [s for s in sheets if not s.skipped]
    results = []

    for sheet in printable:
        # Wrap single sheet in a list for the renderer
        html_content = render_full_document([sheet])
        html_doc = HTML(string=html_content, base_url=_get_base_url())
        pdf_bytes = html_doc.write_pdf()

        # Sanitize filename
        safe_name = sheet.name.replace("/", "_").replace("\\", "_")
        filename = f"{safe_name}.pdf"
        results.append((filename, pdf_bytes))

    return results
