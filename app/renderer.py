"""
HTML/CSS renderer for schedule pages.

Transforms parsed SheetData into print-ready HTML pages,
each designed to fill exactly one A4 portrait page.
"""

import math
from jinja2 import Environment, FileSystemLoader, select_autoescape
import os

from .parser import SheetData, CellData
from . import config

PT_TO_MM = 0.353
CYRILLIC_CHAR_WIDTH_FACTOR = 0.48  # avg char width as fraction of font size (pt)


def _compute_column_widths_mm(num_cols: int) -> list[float]:
    """Compute column widths in mm."""
    avail_w = config.PAGE_WIDTH_MM - config.MARGIN_LEFT_MM - config.MARGIN_RIGHT_MM

    if num_cols <= 3:
        return [avail_w / num_cols] * num_cols

    num_class_cols = num_cols - 3

    if num_class_cols <= 2:
        day_pct, num_pct, time_pct = 0.09, 0.04, 0.12
    else:
        day_pct, num_pct, time_pct = 0.07, 0.035, 0.10

    remaining = 1.0 - day_pct - num_pct - time_pct
    class_pct = remaining / num_class_cols

    pcts = [day_pct, num_pct, time_pct] + [class_pct] * num_class_cols
    return [avail_w * p for p in pcts]


def _compute_column_widths_pct(num_cols: int) -> list[float]:
    """Compute column widths as CSS percentages."""
    if num_cols <= 3:
        return [100.0 / num_cols] * num_cols

    num_class_cols = num_cols - 3

    if num_class_cols <= 2:
        day_pct, num_pct, time_pct = 9.0, 4.0, 12.0
    else:
        day_pct, num_pct, time_pct = 7.0, 3.5, 10.0

    remaining = 100.0 - day_pct - num_pct - time_pct
    class_pct = remaining / num_class_cols

    return [day_pct, num_pct, time_pct] + [class_pct] * num_class_cols


def _lines_needed(text: str, col_width_mm: float, font_pt: float) -> int:
    """Calculate how many lines a text needs in a column at given font size."""
    if not text:
        return 1

    char_width_mm = font_pt * CYRILLIC_CHAR_WIDTH_FACTOR * PT_TO_MM
    if char_width_mm <= 0:
        return 1

    # Account for cell padding reducing available width
    effective_width = col_width_mm - 2.0  # 1mm padding each side
    if effective_width <= 0:
        effective_width = col_width_mm

    chars_per_line = max(1, int(effective_width / char_width_mm))
    text_len = len(text.strip())

    return max(1, math.ceil(text_len / chars_per_line))


def _simulate_table_height(sheet: SheetData, font_pt: float, col_widths_mm: list[float]) -> float:
    """
    Simulate the total table height in mm for a given font size.

    Accounts for text wrapping in each cell and takes the tallest cell per row.
    """
    font_mm = font_pt * PT_TO_MM
    line_height = 1.15
    cell_padding_mm = 1.6  # total vertical padding per cell
    border_mm = 0.3

    total_height = 0.0

    for row_idx, row in enumerate(sheet.rows):
        if row_idx == 0:
            # Title row — single line, larger font
            row_h = font_mm * 1.3 * line_height + 3.0 + border_mm
        elif row_idx == 1:
            # Teacher row — smaller font
            row_h = font_mm * 0.78 * line_height + 1.0 + border_mm
        elif row_idx == 2:
            # Header row
            row_h = font_mm * line_height + 2.0 + border_mm
        else:
            # Data row — find the tallest cell
            max_lines = 1
            for col_idx, cell in enumerate(row):
                if cell.is_merged_slave or not cell.value:
                    continue
                if cell.merge_height > 1:
                    continue  # Day cells span multiple rows, don't count their height per row

                if col_idx < len(col_widths_mm):
                    col_w = col_widths_mm[col_idx]
                else:
                    col_w = col_widths_mm[-1] if col_widths_mm else 30.0

                # Time column uses smaller font
                if col_idx == 2:
                    lines = _lines_needed(cell.value, col_w, font_pt * 0.78)
                else:
                    effective_font = font_pt * 0.92 if col_idx >= 3 else font_pt
                    lines = _lines_needed(cell.value, col_w, effective_font)

                max_lines = max(max_lines, lines)

            row_h = max_lines * font_mm * line_height + cell_padding_mm + border_mm

        total_height += row_h

    return total_height


def _estimate_optimal_font_size(sheet: SheetData) -> float:
    """
    Find the largest font size where the table fits on one A4 page.

    Uses binary search over font sizes, simulating the table height for each.
    """
    # Use 95% of available height as safety margin — WeasyPrint rendering
    # may differ slightly from our simulation due to sub-pixel rounding
    avail_h = (config.PAGE_HEIGHT_MM - config.MARGIN_TOP_MM - config.MARGIN_BOTTOM_MM) * 0.95
    col_widths_mm = _compute_column_widths_mm(sheet.num_cols)

    if sheet.num_rows == 0 or sheet.num_cols == 0:
        return config.MAX_FONT_SIZE_PT

    # Binary search: find largest font where height <= avail_h
    lo = config.MIN_FONT_SIZE_PT
    hi = config.MAX_FONT_SIZE_PT
    best = lo

    for _ in range(50):
        mid = (lo + hi) / 2.0
        h = _simulate_table_height(sheet, mid, col_widths_mm)

        if h <= avail_h:
            best = mid
            lo = mid
        else:
            hi = mid

        if hi - lo < 0.02:
            break

    # Round DOWN to avoid pushing over the limit
    return math.floor(best * 10) / 10.0


def render_full_document(sheets: list[SheetData]) -> str:
    """Render all schedule sheets into a complete HTML document for PDF export."""
    printable_sheets = [s for s in sheets if not s.skipped]

    if not printable_sheets:
        return "<html><body><p>No printable schedule sheets found.</p></body></html>"

    pages = []
    for sheet in printable_sheets:
        font_size = _estimate_optimal_font_size(sheet)
        col_widths = _compute_column_widths_pct(sheet.num_cols)
        pages.append({
            "sheet": sheet,
            "font_size": font_size,
            "col_widths": col_widths,
        })

    template_dir = os.path.join(os.path.dirname(__file__), "templates")
    env = Environment(
        loader=FileSystemLoader(template_dir),
        autoescape=select_autoescape(["html"]),
    )
    env.filters["enumerate_rows"] = lambda iterable: list(enumerate(iterable))
    template = env.get_template("document.html")

    return template.render(
        pages=pages,
        config=config,
    )
