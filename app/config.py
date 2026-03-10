"""
Configuration for schedule converter.

All configurable parameters are centralized here for easy adjustment.
"""

# --- Sheet filtering ---
# Sheet names that are helper/auxiliary data and should NOT be printed.
# Case-insensitive partial match: if any of these strings appear in the sheet name,
# the sheet is skipped.
SKIP_SHEET_PATTERNS = [
    "предметы",    # Subject color palette — helper data
    "звонки",      # Bell schedule — referenced by formulas, not printed standalone
]

# Alternatively, only include sheets matching these patterns (empty = include all non-skipped).
# If non-empty, only sheets containing at least one of these strings are included.
INCLUDE_SHEET_PATTERNS = []  # e.g., ["класс"] to only include class schedule sheets

# --- Bell schedule sheet name (for formula resolution) ---
BELL_SCHEDULE_SHEET = "Звонки"

# --- Page layout ---
PAGE_WIDTH_MM = 210     # A4 portrait width
PAGE_HEIGHT_MM = 297    # A4 portrait height
MARGIN_TOP_MM = 8
MARGIN_BOTTOM_MM = 8
MARGIN_LEFT_MM = 8
MARGIN_RIGHT_MM = 8

# --- Font sizing ---
# The renderer will auto-scale between these bounds to fill the page.
MIN_FONT_SIZE_PT = 8
MAX_FONT_SIZE_PT = 18
TITLE_FONT_SIZE_SCALE = 1.4   # Title font = base font * this factor
HEADER_FONT_SIZE_SCALE = 1.0  # Header row font scale

# --- Table styling ---
CELL_PADDING_MM = 1.5
BORDER_WIDTH_PX = 1
BORDER_COLOR = "#333333"
HEADER_BG_COLOR = "#e8e8e8"
DAY_BG_COLOR = "#f0f0f0"
TITLE_BG_COLOR = "#ffffff"

# --- Export ---
COMBINED_PDF_FILENAME = "расписание.pdf"
SEPARATE_PDF_PREFIX = "расписание_"

# --- Upload ---
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_FILE_SIZE_MB = 20
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
