"""
Excel timetable parser.

Parses school schedule .xlsx files, resolves bell-time formulas,
filters out non-schedule sheets (Предметы, Звонки), and returns
structured data ready for rendering.

--- CONFIGURABLE RULES (edit here) ---
SKIP_SHEETS  – sheet names to always skip
SKIP_PATTERNS – substrings; if a sheet name contains one, skip it
"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
# CONFIGURABLE: sheet filtering
# ──────────────────────────────────────────────
SKIP_SHEETS: set[str] = {"Предметы", "Звонки"}
SKIP_PATTERNS: list[str] = []  # e.g. ["служебн", "helper"]


def should_skip_sheet(name: str) -> tuple[bool, str]:
    """Return (skip, reason) for a given sheet name."""
    if name in SKIP_SHEETS:
        return True, f"Лист «{name}» — служебный (в списке пропуска)"
    for pat in SKIP_PATTERNS:
        if pat.lower() in name.lower():
            return True, f"Лист «{name}» — совпал шаблон «{pat}»"
    return False, ""


# ──────────────────────────────────────────────
# Data structures
# ──────────────────────────────────────────────
@dataclass
class CellData:
    value: str = ""
    is_merged_origin: bool = False
    rowspan: int = 1
    colspan: int = 1
    bold: bool = False
    align_h: str = "center"
    align_v: str = "middle"


@dataclass
class SheetData:
    name: str  # sheet name (e.g. "5 классы")
    title: str  # row-1 title (e.g. "РАСПИСАНИЕ УРОКОВ — 5 КЛАССЫ")
    rows: list[list[CellData]] = field(default_factory=list)
    num_cols: int = 0
    num_rows: int = 0


@dataclass
class ParseResult:
    filename: str
    sheets: list[SheetData] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


# ──────────────────────────────────────────────
# Bell schedule resolver
# ──────────────────────────────────────────────
_BELL_REF = re.compile(r"^=Звонки!([A-Z])(\d+)$", re.IGNORECASE)


def _load_bell_schedule(wb: openpyxl.Workbook) -> dict[str, str]:
    """
    Read the Звонки sheet and build a map:
      "B2" -> "8.35 - 9.15", "C2" -> "8.00-8.40", ...
    """
    bells: dict[str, str] = {}
    if "Звонки" not in wb.sheetnames:
        return bells
    ws = wb["Звонки"]
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                key = f"{get_column_letter(cell.column)}{cell.row}"
                bells[key] = str(cell.value)
    return bells


def _resolve_value(raw, bells: dict[str, str]) -> str:
    """Resolve a cell value, replacing Звонки formula references."""
    if raw is None:
        return ""
    s = str(raw)
    m = _BELL_REF.match(s)
    if m:
        ref = f"{m.group(1).upper()}{m.group(2)}"
        return bells.get(ref, s)
    return s


# ──────────────────────────────────────────────
# Merged-cell map
# ──────────────────────────────────────────────
def _build_merge_map(ws) -> dict[tuple[int, int], tuple[int, int]]:
    """
    Returns {(row, col): (rowspan, colspan)} for merge origins.
    """
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        merge_map[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
    return merge_map


def _cells_in_merges(ws) -> set[tuple[int, int]]:
    """All (row, col) that are *inside* a merge but NOT the origin."""
    covered: set[tuple[int, int]] = set()
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    covered.add((r, c))
    return covered


# ──────────────────────────────────────────────
# Main parse function
# ──────────────────────────────────────────────
def parse_excel(filepath: str | Path) -> ParseResult:
    """
    Parse an Excel timetable file.
    Returns structured data for every schedule sheet.
    """
    filepath = Path(filepath)
    result = ParseResult(filename=filepath.name)

    try:
        wb = openpyxl.load_workbook(str(filepath), data_only=False)
    except Exception as e:
        result.errors.append(f"Не удалось открыть файл: {e}")
        return result

    bells = _load_bell_schedule(wb)

    for sheet_name in wb.sheetnames:
        skip, reason = should_skip_sheet(sheet_name)
        if skip:
            result.skipped.append((sheet_name, reason))
            continue

        ws = wb[sheet_name]
        if ws.max_row is None or ws.max_row < 3:
            result.skipped.append((sheet_name, "Слишком мало строк"))
            continue

        merge_map = _build_merge_map(ws)
        covered = _cells_in_merges(ws)

        # Determine actual data bounds
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1

        # Read title from row 1
        title_val = _resolve_value(ws.cell(1, 1).value, bells)

        sheet_data = SheetData(
            name=sheet_name,
            title=title_val,
            num_cols=max_col,
            num_rows=max_row,
        )

        for r in range(1, max_row + 1):
            row_cells: list[CellData] = []
            for c in range(1, max_col + 1):
                if (r, c) in covered:
                    # This cell is eaten by a merge — skip in output
                    continue

                cell = ws.cell(r, c)
                raw_val = cell.value if not isinstance(cell, MergedCell) else None

                # For merged origins, get the value from the actual cell
                if isinstance(cell, MergedCell):
                    # shouldn't happen for origins, but be safe
                    val = ""
                else:
                    val = _resolve_value(raw_val, bells)

                cd = CellData(value=val)

                # Merge info
                if (r, c) in merge_map:
                    cd.is_merged_origin = True
                    cd.rowspan, cd.colspan = merge_map[(r, c)]

                # Font style
                if cell.font and cell.font.bold:
                    cd.bold = True

                # Alignment
                if cell.alignment:
                    cd.align_h = cell.alignment.horizontal or "center"
                    cd.align_v = cell.alignment.vertical or "center"

                row_cells.append(cd)

            sheet_data.rows.append(row_cells)

        result.sheets.append(sheet_data)

    wb.close()
    return result
