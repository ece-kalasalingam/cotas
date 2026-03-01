import pandas as pd
"""
============================================================
CO Attainment System - Global Policy & System Constants
============================================================

This module contains ALL system-level constants.

RULE:
No hardcoded values should exist in generator.py,
calculator.py, validator.py, or UI files.

If policy changes, edit this file only.
============================================================
"""

# ==========================================================
# VERSIONING
# ==========================================================

SYSTEM_VERSION = "1.0.0"
REGULATION_VERSION = "R2025"


# ==========================================================
# ATTAINMENT POLICY
# ==========================================================

# Direct vs Indirect Contribution
DIRECT_RATIO: float = 0.8
INDIRECT_RATIO: float = 0.2

# Safety Check (prevents silent configuration mistakes)
if round(DIRECT_RATIO + INDIRECT_RATIO, 5) != 1.0:
    raise ValueError("DIRECT_RATIO + INDIRECT_RATIO must equal 1.0")


# ==========================================================
# EXCEL SHEET CONSTANTS
# ==========================================================

# One unified password for all sheets (template + reports)
WORKBOOK_PASSWORD: str = "admin"

# Protection behavior flags
ALLOW_SORT: bool = True
ALLOW_FILTER: bool = True
ALLOW_SELECT_LOCKED: bool = False
ALLOW_SELECT_UNLOCKED: bool = False

HEADER_PATTERNFILL_COLOR = "F2F2F2"
HEADER_PATTERNFILL_TYPE = "solid"


# ==========================================================
# MARK ENTRY RULES
# ==========================================================

ABSENT_SYMBOL: str = "A"

# Numeric formatting
DEFAULT_NUMBER_FORMAT: str = "0.00"

# Marks validation lower bound
MIN_MARK_VALUE: float = 0.0


# ==========================================================
# INDIRECT TOOL POLICY
# ==========================================================

LIKERT_MIN: int = 1
LIKERT_MAX: int = 5

if LIKERT_MIN >= LIKERT_MAX:
    raise ValueError("LIKERT_MIN must be less than LIKERT_MAX")


# ==========================================================
# PAGE LAYOUT DEFAULTS
# ==========================================================

# Margins (in inches)
MARGIN_LEFT: float = 0.2
MARGIN_RIGHT: float = 0.2
MARGIN_TOP: float = 0.5
MARGIN_BOTTOM: float = 0.5
MARGIN_HEADER: float = 0.2
MARGIN_FOOTER: float = 0.2

# ==========================================================
# SYSTEM SHEET NAMES
# ==========================================================

SYSTEM_HASH_SHEET = "__SYSTEM_HASH__"
COURSE_INFO_SHEET = "Course_Info"


# ==========================================================
# INTERNAL SAFETY LIMITS
# ==========================================================

MAX_EXCEL_SHEETNAME_LENGTH = 31
MAX_CO_LIMIT = 50  # Safety upper bound

# ==========================================================
# ATTAINMENT THRESHOLDS
# ==========================================================

PASS_MARK: float = 40.0
THRESHOLD_MARK: float = 60.0
HIGH_BENCHMARK_MARK: float = 80.0

STATUS_NORMAL = 0
STATUS_A = 1
STATUS_NA = 2

# ==========================================================
# END OF CONSTANTS
# ==========================================================
