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
MAIN_WINDOW_TITLE_TEXT_KEY = "app.main_window_title"
APP_ORGANIZATION = "FOCUS"
MAIN_SPLASH_MS = 1500
SINGLE_INSTANCE_LOCK_TIMEOUT_MS = 0
SINGLE_INSTANCE_CLIENT_CONNECT_TIMEOUT_MS = 500
SINGLE_INSTANCE_CLIENT_WRITE_TIMEOUT_MS = 300
SINGLE_INSTANCE_CLIENT_ACK_TIMEOUT_MS = 700
SINGLE_INSTANCE_SERVER_IO_TIMEOUT_MS = 150
SINGLE_INSTANCE_ACTIVATE_PAYLOAD = b"ACTIVATE"
SINGLE_INSTANCE_ACK_PAYLOAD = b"OK"
UI_STANDARD_TIMEOUT_MS = 3000
STARTUP_TOAST_DURATION_MS = 2200
STARTUP_TOAST_QUIT_DELAY_MS = 2300
QT_ADAPTIVE_STRUCTURE_SENSITIVITY = "1"

SYSTEM_VERSION = "1.0.5"
ID_COURSE_SETUP = "COURSE_SETUP_V2"
UI_LANGUAGE = "en"  # Default UI language policy is explicit English.

# Splash defaults
SPLASH_STATUS_COLOR = "#EAF7F5"
# Main window sizing defaults (optimized for common HD 1366x768 displays, including 125% scaling)
WINDOW_TARGET_HEIGHT_RATIO = 0.95
WINDOW_STANDARD_HEIGHT = 560
WINDOW_HEIGHT_CAP = WINDOW_STANDARD_HEIGHT
WINDOW_ADAPTIVE_HEIGHT_CAP_RATIO = 0.72
WINDOW_ADAPTIVE_HEIGHT_CAP_MAX = 1180
WINDOW_WIDTH_TO_HEIGHT_RATIO = 1.57
WINDOW_MIN_WIDTH = 1005
MAIN_ACTIVITY_ICON_SIZE = 30
STATUS_FLASH_TIMEOUT_MS = UI_STANDARD_TIMEOUT_MS
MAIN_HIDDEN_ACTIVITY_MODULE_KEYS = ("HelpModule", "AboutModule")
OUTPUT_LINK_SEPARATOR = "::"
OUTPUT_LINK_MODE_FOLDER = "folder"
OUTPUT_LINK_MODE_FILE = "file"
MODULE_LEFT_PANE_WIDTH_OFFSET = 116
MODULE_LEFT_PANE_CONTENT_MARGINS = (12, 8, 18, 8)
MODULE_LEFT_PANE_LAYOUT_SPACING = 10
INSTRUCTOR_INFO_TAB_FIXED_HEIGHT = 200
WIN32_SHOW_WINDOW_RESTORE = 9

ABOUT_ICON_SIZE = 72
ABOUT_CONTRIBUTORS_FILE = "about_contributors.txt"
APP_REPOSITORY_URL = "https://github.com/ece-kalasalingam/cotas"
FILE_ITEM_REMOVE_BUTTON_SIZE = 24
FILE_ITEM_REMOVE_BUTTON_ICON_SIZE = 16

# ==========================================================
# ATTAINMENT POLICY
# ==========================================================

# Direct vs Indirect Contribution
DIRECT_RATIO: float = 0.8
INDIRECT_RATIO: float = 0.2
# CO attainment level thresholds (percentage scale: 0-100)
LEVEL_1_THRESHOLD: float = 40.0
LEVEL_2_THRESHOLD: float = 60.0
LEVEL_3_THRESHOLD: float = 75.0
CO_ATTAINMENT_PERCENT_DEFAULT: float = 80.0
CO_ATTAINMENT_LEVEL_DEFAULT: int = 2


# ==========================================================
# EXCEL SHEET CONSTANTS
# ==========================================================

# Protection behavior flags
ALLOW_SORT: bool = True
ALLOW_FILTER: bool = True
ALLOW_SELECT_LOCKED: bool = False
ALLOW_SELECT_UNLOCKED: bool = True

# ==========================================================
# MARK ENTRY RULES
# ==========================================================

# Marks validation lower bound
MIN_MARK_VALUE: float = 0.0
MARKS_ENTRY_VALIDATION_ERROR_TITLE: str = "Invalid marks"
MARKS_ENTRY_VALIDATION_ERROR_RANGE_TEMPLATE: str = (
    "Enter A/a or a numeric mark between {minimum} and {maximum}."
)
MARKS_ENTRY_INDIRECT_VALIDATION_ERROR_RANGE_TEMPLATE: str = (
    "Enter A/a or a numeric Likert value between {minimum} and {maximum}."
)
MARKS_ENTRY_ROW_HEADERS = ("#", "Reg. No.", "Student Name")
MARKS_ENTRY_TOTAL_LABEL = "Total"
MARKS_ENTRY_QUESTION_PREFIX = "Q"
MARKS_ENTRY_CO_MARKS_LABEL_PREFIX = "Marks for CO"

# Final CO report (step 3 generate)
CO_REPORT_DIRECT_SHEET_SUFFIX = "_Direct"
CO_REPORT_INDIRECT_SHEET_SUFFIX = "_Indirect"
CO_REPORT_ABSENT_TOKEN = "A"
CO_REPORT_NOT_APPLICABLE_TOKEN = "NA"
CO_REPORT_HEADER_SERIAL = "#"
CO_REPORT_HEADER_REG_NO = "Reg. No."
CO_REPORT_HEADER_STUDENT_NAME = "Student Name"
CO_REPORT_HEADER_TOTAL = "Total"
CO_REPORT_HEADER_TOTAL_100 = "Total (100%)"
CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE = "Total ({ratio}%)"
CO_REPORT_PERCENT_SYMBOL = "%"
CO_REPORT_MAX_DECIMAL_PLACES = 2
CO_REPORT_SCALED_LABEL_TEMPLATE = "scaled 0-{max_value}"
COMPONENT_NAME_LABEL = "Component name"
CO_LABEL = "CO"
INSTRUCTOR_MAX_LABEL = "Max."
WORKBOOK_TEMP_SUFFIX = ".tmp"
WORKFLOW_STEP_TIMEOUT_ENV_VAR = "FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS"
WORKFLOW_OPERATION_GENERATE_COURSE_DETAILS_TEMPLATE = "generate_course_details_template"
WORKFLOW_OPERATION_VALIDATE_COURSE_DETAILS_WORKBOOK = "validate_course_details_workbook"
WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE = "generate_marks_template"
CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK = "generating co analysis workbook"
WORKFLOW_PAYLOAD_KEY_TEMPLATE_ID = "template_id"
WORKFLOW_PAYLOAD_KEY_OUTPUT = "output"
WORKFLOW_PAYLOAD_KEY_SOURCE = "source"
WORKFLOW_PAYLOAD_KEY_PATH = "path"
LOG_EXTRA_KEY_USER_MESSAGE = "user_message"
LOG_EXTRA_KEY_JOB_ID = "job_id"
LOG_EXTRA_KEY_STEP_ID = "step_id"
PROCESS_MESSAGE_SUCCESS_SUFFIX = " completed successfully."
PROCESS_MESSAGE_CANCELLED_TEMPLATE = "%s cancelled by user/system request."
WORKFLOW_USER_MESSAGE_STARTED_SUFFIX = " started."
WORKFLOW_USER_MESSAGE_COMPLETED_TEMPLATE = " completed in {duration_ms} ms."
WORKFLOW_USER_MESSAGE_CANCELLED_TEMPLATE = " cancelled after {duration_ms} ms."
WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE = " failed after {duration_ms} ms."
WORKFLOW_TIMEOUT_ERROR_TEMPLATE = "{operation} exceeded timeout of {timeout_seconds} seconds."
LAYOUT_MANIFEST_KEY_SHEET_ORDER = "sheet_order"
LAYOUT_MANIFEST_KEY_SHEETS = "sheets"
LAYOUT_SHEET_SPEC_KEY_NAME = "name"
LAYOUT_SHEET_SPEC_KEY_KIND = "kind"
LAYOUT_SHEET_SPEC_KEY_HEADER_ROW = "header_row"
LAYOUT_SHEET_SPEC_KEY_HEADERS = "headers"
LAYOUT_SHEET_SPEC_KEY_ANCHORS = "anchors"
LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS = "formula_anchors"
LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT = "student_count"
LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH = "student_identity_hash"
LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE = "mark_structure"
LAYOUT_SHEET_KIND_DIRECT_CO_WISE = "direct_co_wise"
LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE = "direct_non_co_wise"
LAYOUT_SHEET_KIND_INDIRECT = "indirect"
WORKBOOK_INTEGRITY_SCHEMA_VERSION = 1


# ==========================================================
# INDIRECT TOOL POLICY
# ==========================================================

LIKERT_MIN: int = 1
LIKERT_MAX: int = 5


# ==========================================================
# PAGE LAYOUT DEFAULTS
# ==========================================================

# ==========================================================
# SYSTEM SHEET NAMES
# ==========================================================

ASSESSMENT_VALIDATION_LAST_ROW = 300
COURSE_METADATA_FACULTY_NAME_KEY = "faculty_name"
CO_REPORT_METADATA_OUTCOME_FIELD = "Course Outcome"
CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE = "CO{co} - Direct"
CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE = "CO{co} - Indirect"
CO_ANALYSIS_SHEET_FOOTER_TEXT = "Ssystem generated report"
CO_ANALYSIS_INSTITUTION_ROW_1 = "KALASALINGAM ACADEMY OF RESEARCH AND EDUCATION"
CO_ANALYSIS_INSTITUTION_ROW_2 = "(Deemed to be University)"
CO_ANALYSIS_INSTITUTION_ROW_3 = "Krishnankoil Campus"
CO_ANALYSIS_INSTITUTION_ROW_4 = "Department of Electronics and Communication Engineering"
CO_ANALYSIS_INSTITUTION_ROW_5 = "B.Tech. Electronics and Communication Engineering"
CO_ANALYSIS_INSTITUTION_ROWS = (
    CO_ANALYSIS_INSTITUTION_ROW_1,
    CO_ANALYSIS_INSTITUTION_ROW_2,
    CO_ANALYSIS_INSTITUTION_ROW_3,
    CO_ANALYSIS_INSTITUTION_ROW_4,
    CO_ANALYSIS_INSTITUTION_ROW_5,
)
MARKS_TEMPLATE_NAME_SUFFIX = "Marks"
CO_REPORT_TEMPLATE_NAME_SUFFIX = "COReport"
FILENAME_JOIN_SEPARATOR = "_"
FILE_EXTENSION_XLSX = ".xlsx"


# ==========================================================
# INTERNAL SAFETY LIMITS
# ==========================================================

MAX_EXCEL_SHEETNAME_LENGTH = 31

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
