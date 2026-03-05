import os

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
APP_EXECUTABLE_NAME = "focus"
APP_INTERNAL_NAME = "focus"
APP_PRODUCT_NAME = "Focus"
APP_SUBTITLE = "Framework for Outcome Computation and Unification System"
APP_SUBTITLE_TEXT_KEY = "app.subtitle"
MAIN_WINDOW_TITLE_TEXT_KEY = "app.main_window_title"
APP_ORGANIZATION = "FOCUS"
MAIN_SPLASH_MS = 1500
SINGLE_INSTANCE_LOCK_TIMEOUT_MS = 0
THEME_REFRESH_DEBOUNCE_MS = 120
THEME_SETUP_DEFER_MS = 0
UI_STANDARD_TIMEOUT_MS = 3000
STARTUP_TOAST_DURATION_MS = 2200
STARTUP_TOAST_QUIT_DELAY_MS = 2300
QT_ADAPTIVE_STRUCTURE_SENSITIVITY = "1"

SYSTEM_VERSION = "1.0.0"
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
WINDOW_STANDARD_HEIGHT = 640
WINDOW_HEIGHT_CAP = WINDOW_STANDARD_HEIGHT
WINDOW_WIDTH_TO_HEIGHT_RATIO = 1.57
WINDOW_MIN_WIDTH = 1005
WINDOW_MIN_HEIGHT = WINDOW_STANDARD_HEIGHT
MAIN_ACTIVITY_ICON_SIZE = 30
STATUS_FLASH_TIMEOUT_MS = UI_STANDARD_TIMEOUT_MS
MAIN_WINDOW_CONTENT_MARGINS = (0, 0, 0, 0)

ABOUT_ICON_SIZE = 72
ABOUT_LAYOUT_MARGIN = 40
ABOUT_LAYOUT_SPACING = 18
ABOUT_HEADER_SPACING = 20
ABOUT_TITLE_SPACING = 4
ABOUT_BODY_GAP_LARGE = 8
ABOUT_BODY_GAP_SMALL = 4
INSTRUCTOR_RAIL_TITLE_FONT_SIZE = 14
INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE = 18
INSTRUCTOR_STEP_LIST_SPACING = 2
INSTRUCTOR_RAIL_MAX_WIDTH = 290
INSTRUCTOR_CARD_MARGIN = 20
INSTRUCTOR_CARD_SPACING = 14
INSTRUCTOR_STEP2_ACTION_SPACING = 10
INSTRUCTOR_STEP2_ACTION_MARGIN = 0
INSTRUCTOR_TOP_LAYOUT_MARGINS = (0, 0, 0, 0)
INSTRUCTOR_INFO_TAB_FIXED_HEIGHT = 220
INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS = (0, 0, 0, 0)
INSTRUCTOR_INFO_TAB_LAYOUT_SPACING = 0
HELP_LAYOUT_CONTENT_MARGINS = (0, 0, 0, 0)

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
QPushButton {
    padding: 6px 12px;
    min-width: 150px;
    min-height: 30px;
    border-radius: 6px;
    border: 1px solid palette(mid);
}
QTabWidget#instructorInfoTabs::pane {
    border: none;
    background: palette(base);
}
QTabWidget#instructorInfoTabs QTabBar::tab:first {
    margin-left: 8px;
}
QTabWidget#instructorInfoTabs QPlainTextEdit,
QTabWidget#instructorInfoTabs QTextBrowser {
    border: 1px solid palette(mid);
    border-radius: 8px;
    background: palette(base);
    padding: 8px;
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

# One unified password for all sheets (template + reports).
# This must always be provided via environment variable in all environments.
WORKBOOK_PASSWORD_ENV_VAR = "FOCUS_WORKBOOK_PASSWORD"
WORKBOOK_PASSWORD: str = os.getenv(WORKBOOK_PASSWORD_ENV_VAR, "").strip()
if not WORKBOOK_PASSWORD:
    raise ConfigurationError(
        f"{WORKBOOK_PASSWORD_ENV_VAR} is required and must not be empty"
    )
if len(WORKBOOK_PASSWORD) < 12:
    raise ConfigurationError(
        f"{WORKBOOK_PASSWORD_ENV_VAR} must be at least 12 characters long"
    )
WORKBOOK_PASSWORD_PREVIOUS_ENV_VAR = "FOCUS_WORKBOOK_PASSWORD_PREVIOUS"
WORKBOOK_PASSWORD_PREVIOUS: tuple[str, ...] = tuple(
    secret.strip()
    for secret in os.getenv(WORKBOOK_PASSWORD_PREVIOUS_ENV_VAR, "").split(",")
    if secret.strip()
)
for _previous_secret in WORKBOOK_PASSWORD_PREVIOUS:
    if len(_previous_secret) < 12:
        raise ConfigurationError(
            f"{WORKBOOK_PASSWORD_PREVIOUS_ENV_VAR} entries must be at least 12 characters long"
        )
WORKBOOK_SIGNATURE_VERSION_ENV_VAR = "FOCUS_WORKBOOK_SIGNATURE_VERSION"
WORKBOOK_SIGNATURE_VERSION = os.getenv(WORKBOOK_SIGNATURE_VERSION_ENV_VAR, "v1").strip().lower() or "v1"
if WORKBOOK_SIGNATURE_VERSION not in {"v1"}:
    raise ConfigurationError(
        f"{WORKBOOK_SIGNATURE_VERSION_ENV_VAR} must be one of: v1"
    )

# Protection behavior flags
ALLOW_SORT: bool = True
ALLOW_FILTER: bool = True
ALLOW_SELECT_LOCKED: bool = False
ALLOW_SELECT_UNLOCKED: bool = True

HEADER_PATTERNFILL_COLOR = "F2F2F2"


# ==========================================================
# MARK ENTRY RULES
# ==========================================================

# Marks validation lower bound
MIN_MARK_VALUE: float = 0.0
MARKS_ENTRY_VALIDATION_FORMULA: str = '=OR(D4="A",D4="a",AND(ISNUMBER(D4),D4>=0,D4<=D$3))'
MARKS_ENTRY_VALIDATION_ERROR_TITLE: str = "Invalid marks"
MARKS_ENTRY_VALIDATION_ERROR_MESSAGE: str = "Enter A/a or a numeric mark within allowed maximum."
MARKS_ENTRY_INDIRECT_VALIDATION_ERROR_MESSAGE: str = (
    "Enter A/a or a numeric Likert value within allowed range."
)
MARKS_ENTRY_ROW_HEADERS = ("#", "Reg. No.", "Student Name")
MARKS_ENTRY_TOTAL_LABEL = "Total"
MARKS_ENTRY_CO_PREFIX = "CO"
MARKS_ENTRY_QUESTION_PREFIX = "Q"
MARKS_ENTRY_CO_MARKS_LABEL_PREFIX = "Marks for CO"
MARKS_ENTRY_COS_LABEL = "COs"


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

# ==========================================================
# SYSTEM SHEET NAMES
# ==========================================================

SYSTEM_HASH_SHEET = "__SYSTEM_HASH__"
SYSTEM_LAYOUT_SHEET = "__SYSTEM_LAYOUT__"
COURSE_METADATA_SHEET = "Course_Metadata"
ASSESSMENT_CONFIG_SHEET = "Assessment_Config"
QUESTION_MAP_SHEET = "Question_Map"
STUDENTS_SHEET = "Students"

COURSE_METADATA_HEADERS = ("Field", "Value")
ASSESSMENT_CONFIG_HEADERS = ("Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct")
QUESTION_MAP_HEADERS = ("Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO")
STUDENTS_HEADERS = ("Reg_No", "Student_Name")

ASSESSMENT_VALIDATION_YES_NO_OPTIONS = ("YES", "NO")
ASSESSMENT_VALIDATION_INPUT_TITLE = "Direct"
ASSESSMENT_VALIDATION_INPUT_MESSAGE = "Select YES or NO"
ASSESSMENT_VALIDATION_LAST_ROW = 300

SYSTEM_HASH_TEMPLATE_ID_HEADER = "Template_ID"
SYSTEM_HASH_TEMPLATE_HASH_HEADER = "Template_Hash"
SYSTEM_HASH_TEMPLATE_ID_KEY = "template_id"
SYSTEM_HASH_TEMPLATE_HASH_KEY = "template_hash"
SYSTEM_LAYOUT_MANIFEST_HEADER = "Layout_Manifest"
SYSTEM_LAYOUT_MANIFEST_HASH_HEADER = "Layout_Hash"
SYSTEM_LAYOUT_MANIFEST_KEY = "layout_manifest"
SYSTEM_LAYOUT_MANIFEST_HASH_KEY = "layout_hash"
COURSE_METADATA_TOTAL_OUTCOMES_KEY = "total_outcomes"


# ==========================================================
# INTERNAL SAFETY LIMITS
# ==========================================================

MAX_EXCEL_SHEETNAME_LENGTH = 31
WEIGHT_TOTAL_EXPECTED = 100.0
WEIGHT_TOTAL_ROUND_DIGITS = 6

# Toast defaults
TOAST_DEFAULT_DURATION_MS = UI_STANDARD_TIMEOUT_MS
TOAST_ERROR_DURATION_MS = 4500
TOAST_MARGIN = 16
TOAST_CONTENT_MARGIN_X = 12
TOAST_CONTENT_MARGIN_Y = 8
TOAST_CONTENT_MARGIN_LEFT = TOAST_CONTENT_MARGIN_X
TOAST_CONTENT_MARGIN_TOP = TOAST_CONTENT_MARGIN_Y
TOAST_CONTENT_MARGIN_RIGHT = TOAST_CONTENT_MARGIN_X
TOAST_CONTENT_MARGIN_BOTTOM = TOAST_CONTENT_MARGIN_Y
TOAST_CONTENT_SPACING = 2
TOAST_SHADOW_BLUR_RADIUS = 24
TOAST_SHADOW_OFFSET_X = 0
TOAST_SHADOW_OFFSET_Y = 6
TOAST_SHADOW_ALPHA = 60
TOAST_MAX_WIDTH = 460

# ==========================================================
# END OF CONSTANTS
# ==========================================================
