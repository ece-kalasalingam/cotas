from common.exceptions import ConfigurationError

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

APP_NAME = "FOCUS"
APP_SUBTITLE = "Framework for Outcome Computation and Unification System"
MAIN_WINDOW_TITLE = f"{APP_NAME} - {APP_SUBTITLE}"
APP_ORGANIZATION = APP_NAME
MAIN_SPLASH_MS = 1500

SYSTEM_VERSION = "1.0.0"
REGULATION_VERSION = "R2025"
ID_COURSE_SETUP = "COURSE_SETUP_V1"
UI_LANGUAGE = "auto"  # "auto" uses OS user language; fallback is English (1033).
UI_FONT_FAMILY = "Segoe UI"

# Splash defaults
SPLASH_WIDTH = 520
SPLASH_HEIGHT = 240
SPLASH_BG_COLOR = "#2957A4"
SPLASH_TEXT_COLOR = "#ffffff"
SPLASH_STATUS_COLOR = "#EAF7F5"
SPLASH_TITLE_FONT_SIZE = 20

# Main window sizing defaults
WINDOW_TARGET_HEIGHT_RATIO = 0.8
WINDOW_HEIGHT_CAP = 640
WINDOW_WIDTH_TO_HEIGHT_RATIO = 1.57
WINDOW_MIN_WIDTH = 1005
WINDOW_MIN_HEIGHT = 640
MAIN_ACTIVITY_ICON_SIZE = 30

ABOUT_ICON_SIZE = 72
INSTRUCTOR_RAIL_TITLE_FONT_SIZE = 14
INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE = 18

# Shared UI style snippets
ABOUT_APP_NAME_STYLE = "font-size: 24px; font-weight: 600;"
ABOUT_SUBTITLE_STYLE = "font-size: 15px;"
ABOUT_META_STYLE = "font-size: 12px; color: gray;"
ABOUT_BODY_STYLE = "font-size: 12px;"
ABOUT_COPYRIGHT_STYLE = "font-size: 11px; color: gray;"

MAIN_ACTIVITYBAR_STYLESHEET = """
QToolBar {
    spacing: 0px;
    padding: 5px;
}
QToolButton {
    min-width: 80px;
    max-width: 80px;
    padding: 5px;
    border: none;
    border-radius: 4px;
}
QToolButton:hover {
    background-color: rgba(255, 255, 255, 0.1);
}
QToolButton:checked {
    background-color: rgba(22, 160, 133, 0.2);
    border-bottom: 2px solid #16A085;
}
"""

INSTRUCTOR_PANEL_STYLESHEET = """
QFrame#stepRail {
    border: 1px solid palette(mid);
    border-radius: 12px;
    background-color: palette(base);
}
QFrame#activeCard {
}
QListWidget#stepList {
    outline: none;
    background-color: transparent;
}
QListWidget#stepList::item {
    padding: 8px 8px;
}
QListWidget#stepList::item:selected,
QListWidget#stepList::item:selected:!active {
    border-left: 4px solid palette(highlight);
}
QPushButton#primaryAction {
    padding: 6px 12px;
    min-width: 150px;
    min-height: 30px;
    border-radius: 6px;
}
QPushButton#primaryAction:enabled {
    background-color: palette(highlight);
    color: palette(highlighted-text);
    border: none;
}
"""


# ==========================================================
# ATTAINMENT POLICY
# ==========================================================

# Direct vs Indirect Contribution
DIRECT_RATIO: float = 0.8
INDIRECT_RATIO: float = 0.2

# Safety Check (prevents silent configuration mistakes)
if round(DIRECT_RATIO + INDIRECT_RATIO, 5) != 1.0:
    raise ConfigurationError("DIRECT_RATIO + INDIRECT_RATIO must equal 1.0")


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
    raise ConfigurationError("LIKERT_MIN must be less than LIKERT_MAX")


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
