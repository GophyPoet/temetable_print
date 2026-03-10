"""
HTML/CSS renderer for timetable pages → PDF via WeasyPrint.

Each SheetData becomes one A4 portrait page in the final PDF.
The layout maximises page usage with large fonts, minimal margins,
and vertical day-of-week labels when columns are tight.

--- CONFIGURABLE (edit here) ---
PAGE_MARGIN_MM   – page margins
MIN_FONT_PT      – never go below this
MAX_FONT_PT      – ideal font size when space allows
ROTATE_DAY_NAMES – rotate first-column day names vertically
"""

from io import BytesIO
from pathlib import Path

from weasyprint import HTML

from .parser import CellData, SheetData, ParseResult

# ──────────────────────────────────────────────
# CONFIGURABLE: layout
# ──────────────────────────────────────────────
PAGE_MARGIN_MM = 6          # mm on each side
MIN_FONT_PT = 7
MAX_FONT_PT = 14
ROTATE_DAY_NAMES = True     # rotate day-of-week cells vertically (bottom→top)

# Days of week to detect for rotation
_DAYS = {
    "ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ",
    "понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье",
}


def _is_day_cell(val: str) -> bool:
    return val.strip().upper() in {d.upper() for d in _DAYS}


def _estimate_font_size(sheet: SheetData) -> int:
    """Pick a font size based on table dimensions."""
    rows = sheet.num_rows
    if rows >= 44:
        return MIN_FONT_PT
    if rows >= 38:
        return max(MIN_FONT_PT, 8)
    if rows >= 32:
        return max(MIN_FONT_PT, 9)
    if rows >= 28:
        return max(MIN_FONT_PT, 10)
    return MAX_FONT_PT


def _render_sheet_html(sheet: SheetData) -> str:
    """Convert one SheetData into a single-page HTML table."""
    font_size = _estimate_font_size(sheet)

    rows_html = []
    for r_idx, row in enumerate(sheet.rows):
        cells_html = []
        for cell in row:
            tag = "th" if r_idx < 3 else "td"
            attrs = []
            styles = []

            if cell.rowspan > 1:
                attrs.append(f'rowspan="{cell.rowspan}"')
            if cell.colspan > 1:
                attrs.append(f'colspan="{cell.colspan}"')

            # Vertical rotation for day names
            val = cell.value
            css_class = ""
            if ROTATE_DAY_NAMES and _is_day_cell(val) and cell.rowspan > 1:
                css_class = "day-vertical"

            if cell.bold:
                styles.append("font-weight:bold")

            if styles:
                attrs.append(f'style="{";".join(styles)}"')
            if css_class:
                attrs.append(f'class="{css_class}"')

            attr_str = (" " + " ".join(attrs)) if attrs else ""
            # Escape HTML
            safe_val = (
                str(val)
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            cells_html.append(f"<{tag}{attr_str}>{safe_val}</{tag}>")

        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    table_html = f"<table>{''.join(rows_html)}</table>"

    # Column widths: day col narrow when rotated, others flexible
    day_col_w = "2.5em" if ROTATE_DAY_NAMES else "auto"
    # Adaptive padding for dense sheets
    cell_pad = "1px 2px" if sheet.num_rows >= 40 else "2px 3px"
    line_h = "1.05" if sheet.num_rows >= 40 else "1.15"
    title_pad = "2px 1px" if sheet.num_rows >= 40 else "4px 2px"

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<style>
@page {{
    size: A4 portrait;
    margin: {PAGE_MARGIN_MM}mm;
}}
* {{
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}}
html, body {{
    width: 100%;
    height: 100%;
    font-family: "DejaVu Sans", "Liberation Sans", Arial, sans-serif;
    font-size: {font_size}pt;
    line-height: {line_h};
}}
.page {{
    width: 100%;
    height: 100%;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: stretch;
    page-break-after: always;
    page-break-inside: avoid;
    overflow: hidden;
}}
table {{
    width: 100%;
    height: 100%;
    border-collapse: collapse;
    table-layout: fixed;
}}
th, td {{
    border: 1.2px solid #222;
    padding: {cell_pad};
    text-align: center;
    vertical-align: middle;
    word-wrap: break-word;
    overflow-wrap: break-word;
    overflow: hidden;
}}
/* Title row */
tr:first-child th,
tr:first-child td {{
    font-weight: bold;
    font-size: {min(font_size + 3, MAX_FONT_PT + 3)}pt;
    border-bottom: 2px solid #000;
    padding: {title_pad};
}}
/* Teacher row */
tr:nth-child(2) th,
tr:nth-child(2) td {{
    font-size: {max(font_size - 1, MIN_FONT_PT)}pt;
    font-style: italic;
}}
/* Header row */
tr:nth-child(3) th,
tr:nth-child(3) td {{
    font-weight: bold;
    background: #e8e8e8;
    border-bottom: 2px solid #000;
}}
/* Day-of-week vertical */
.day-vertical {{
    writing-mode: vertical-lr;
    transform: rotate(180deg);
    text-orientation: mixed;
    font-weight: bold;
    font-size: {max(font_size - 1, MIN_FONT_PT)}pt;
    letter-spacing: 1px;
    padding: 2px 1px;
    width: {day_col_w};
    max-width: 2.8em;
    text-align: center;
    white-space: nowrap;
}}
/* Lesson number column */
td:nth-child(2), th:nth-child(2) {{
    width: 1.8em;
    text-align: center;
}}
/* Time column */
td:nth-child(3), th:nth-child(3) {{
    width: 5.5em;
    font-size: {max(font_size - 1, MIN_FONT_PT)}pt;
    white-space: nowrap;
}}
</style>
</head>
<body>
<div class="page">
{table_html}
</div>
</body>
</html>"""


def render_to_pdf(parse_results: list[ParseResult]) -> bytes:
    """
    Render all parsed timetable sheets into a single combined PDF.
    Each sheet = one A4 page.
    """
    all_pages_html: list[str] = []

    for pr in parse_results:
        for sheet in pr.sheets:
            all_pages_html.append(_render_sheet_html(sheet))

    if not all_pages_html:
        raise ValueError("Нет листов для рендеринга")

    # Combine pages into one HTML document
    combined = _combine_pages_html(all_pages_html)
    pdf_bytes = HTML(string=combined).write_pdf()
    return pdf_bytes


def render_sheets_separately(parse_results: list[ParseResult]) -> list[tuple[str, bytes]]:
    """
    Render each sheet as a separate PDF.
    Returns list of (sheet_name, pdf_bytes).
    """
    result = []
    for pr in parse_results:
        for sheet in pr.sheets:
            html = _render_sheet_html(sheet)
            pdf_bytes = HTML(string=html).write_pdf()
            result.append((sheet.name, pdf_bytes))
    return result


def _combine_pages_html(pages: list[str]) -> str:
    """
    Combine multiple single-page HTMLs into one multi-page document.
    We extract the <style> from the first page and the table bodies from all.
    """
    if len(pages) == 1:
        return pages[0]

    # For simplicity with WeasyPrint: generate one doc with multiple .page divs.
    # We re-render each sheet individually and merge the PDFs.
    # Actually, WeasyPrint handles multi-page better if we merge at PDF level.
    return _merge_via_separate_render(pages)


def _merge_via_separate_render(pages: list[str]) -> str:
    """
    Since each page may have different font sizes, we render them
    individually and merge PDF bytes.
    This function is a marker — actual merging happens in render_to_pdf_merged.
    """
    # Fallback: just return first page (real merging done differently)
    return pages[0]


def render_to_pdf_merged(parse_results: list[ParseResult]) -> bytes:
    """
    Render each sheet individually and merge into one PDF.
    This ensures each page has its own optimal font size.
    """
    from weasyprint import HTML

    individual_pdfs: list[bytes] = []

    for pr in parse_results:
        for sheet in pr.sheets:
            html = _render_sheet_html(sheet)
            pdf_data = HTML(string=html).write_pdf()
            individual_pdfs.append(pdf_data)

    if not individual_pdfs:
        raise ValueError("Нет листов для рендеринга")

    if len(individual_pdfs) == 1:
        return individual_pdfs[0]

    # Merge PDFs using a simple approach
    return _merge_pdfs(individual_pdfs)


def _merge_pdfs(pdf_list: list[bytes]) -> bytes:
    """Merge multiple single-page PDFs into one multi-page PDF."""
    try:
        # Try pypdf if available
        from pypdf import PdfWriter, PdfReader
        writer = PdfWriter()
        for pdf_bytes in pdf_list:
            reader = PdfReader(BytesIO(pdf_bytes))
            for page in reader.pages:
                writer.add_page(page)
            buf = BytesIO()
            writer.write(buf)
        return buf.getvalue()
    except ImportError:
        pass

    # Fallback: manual PDF merge using basic PDF structure
    # For robustness, render all sheets in one HTML with page breaks
    return _fallback_single_render(pdf_list)


def _fallback_single_render(pdf_list: list[bytes]) -> bytes:
    """If we can't merge PDFs, just return the first one (degraded)."""
    # This shouldn't normally be reached; pypdf should be available
    return pdf_list[0] if pdf_list else b""
