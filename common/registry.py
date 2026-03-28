from __future__ import annotations

from common.constants import (
    CO_LABEL,
    MARKS_ENTRY_CO_MARKS_LABEL_PREFIX,
    MARKS_ENTRY_QUESTION_PREFIX,
    MARKS_ENTRY_ROW_HEADERS,
    MARKS_ENTRY_TOTAL_LABEL,
)
from common.exceptions import ConfigurationError
from common.sheet_schema import SheetSchema, ValidationRule, WorkbookBlueprint

# Canonical logical keys for COURSE_SETUP workbook sheets.
COURSE_SETUP_SHEET_KEY_COURSE_METADATA = "course_metadata"
COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG = "assessment_config"
COURSE_SETUP_SHEET_KEY_QUESTION_MAP = "question_map"
COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION = "co_description"
COURSE_SETUP_SHEET_KEY_STUDENTS = "students"
CO_REPORT_SHEET_KEY_CO_INDIRECT = "co_indirect"
COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE = "marks_direct_co_wise"
COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE = "marks_direct_non_co_wise"
COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT = "marks_indirect"

SYSTEM_HASH_SHEET_NAME = "__SYSTEM_HASH__"
SYSTEM_HASH_HEADER_TEMPLATE_ID = "Template_ID"
SYSTEM_HASH_HEADER_TEMPLATE_HASH = "Template_Hash"
SYSTEM_HASH_KEY_TEMPLATE_ID = "template_id"
SYSTEM_HASH_KEY_TEMPLATE_HASH = "template_hash"

COURSE_SETUP_ASSESSMENT_VALIDATION_YES_NO_OPTIONS = ("YES", "NO")
COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_TITLE = "Direct"
COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_MESSAGE = "Select YES or NO"
COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS = ("FORMATIVE", "SUMMATIVE")
COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS = (
    "THEORY_EXAM",
    "LAB_WORK",
    "PROJECT",
    "VIVA",
    "PRACTICAL_EXAM",
    "SURVEY",
)
COURSE_SETUP_ASSESSMENT_MODE_OPTIONS = (
    "WRITTEN",
    "HANDS_ON",
    "ORAL",
    "PRESENTATION",
    "WRITTEN+ORAL",
    "HANDS_ON+WRITTEN",
    "HANDS_ON+ORAL",
    "HANDS_ON+WRITTEN+ORAL",
    "PRESENTATION+ORAL",
    "WRITTEN+PRESENTATION",
)
COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS = (
    "INDIVIDUAL",
    "GROUP",
    "INDIVIDUAL+GROUP",
)
COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS = (
    "REMEMBER",
    "UNDERSTAND",
    "APPLY",
    "ANALYZE",
    "EVALUATE",
    "CREATE",
    "SKILL_DEVELOPMENT",
    "MULTIPLE_LEVELS",
)
COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH = 100
COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH = 200
COURSE_METADATA_TOTAL_OUTCOMES_KEY = "total_outcomes"
COURSE_METADATA_COURSE_CODE_KEY = "course_code"
COURSE_METADATA_SEMESTER_KEY = "semester"
COURSE_METADATA_SECTION_KEY = "section"
COURSE_METADATA_ACADEMIC_YEAR_KEY = "academic_year"
COURSE_METADATA_TOTAL_STUDENTS_KEY = "total_students"
WEIGHT_TOTAL_EXPECTED = 100.0
WEIGHT_TOTAL_ROUND_DIGITS = 6

SETUP_STYLE_REGISTRY_V1 = {
    "header": {
        "bold": True,
        "bg_color": "#cccccc",
        "border": 0,
        "align": "center",
        "valign": "vcenter",
    },
    "body": {
        "locked": False,
        "border": 1,
    },
}
SETUP_STYLE_REGISTRY_V2 = {
    "header": {
        "bold": True,
        "bg_color": "#D9EAD3",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    },
    "body": {
        "locked": False,
        "border": 1,
    },
}

_CO_INDIRECT_HEADER_RESOLVER = "course_setup.co_indirect_headers"
_MARKS_DIRECT_CO_WISE_HEADER_RESOLVER = "course_setup.marks_direct_co_wise_headers"
_MARKS_DIRECT_NON_CO_WISE_HEADER_RESOLVER = "course_setup.marks_direct_non_co_wise_headers"
_MARKS_INDIRECT_HEADER_RESOLVER = "course_setup.marks_indirect_headers"


def _rule(
    *,
    first_row: int,
    first_col: int,
    last_col: int,
    options: dict[str, object],
) -> ValidationRule:
    from common.constants import ASSESSMENT_VALIDATION_LAST_ROW

    return ValidationRule(
        first_row=first_row,
        first_col=first_col,
        last_row=ASSESSMENT_VALIDATION_LAST_ROW,
        last_col=last_col,
        options=dict(options),
    )


def _sheet(
    *,
    key: str,
    name: str,
    headers: tuple[str, ...],
    validations: list[ValidationRule] | None = None,
    is_protected: bool = False,
    header_kind: str = "fixed",
    header_resolver: str | None = None,
    header_context: dict[str, object] | None = None,
    sheet_rules: dict[str, object] | None = None,
) -> SheetSchema:
    return SheetSchema(
        key=key,
        name=name,
        header_matrix=[list(headers)],
        header_kind=header_kind,
        header_resolver=header_resolver,
        header_context=dict(header_context or {}),
        validations=list(validations or []),
        is_protected=is_protected,
        sheet_rules=dict(sheet_rules or {}),
    )


def _clone_sheet(schema: SheetSchema) -> SheetSchema:
    return SheetSchema(
        key=schema.key,
        name=schema.name,
        header_matrix=[list(row) for row in schema.header_matrix],
        header_kind=schema.header_kind,
        header_resolver=schema.header_resolver,
        header_context=dict(schema.header_context),
        validations=[
            ValidationRule(
                first_row=rule.first_row,
                first_col=rule.first_col,
                last_row=rule.last_row,
                last_col=rule.last_col,
                options=dict(rule.options),
            )
            for rule in schema.validations
        ],
        is_protected=schema.is_protected,
        sheet_rules=dict(schema.sheet_rules),
    )


def _course_setup_sheet_catalog() -> dict[str, SheetSchema]:
    return {
        COURSE_SETUP_SHEET_KEY_COURSE_METADATA: _sheet(
            key=COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
            name="Course_Metadata",
            headers=("Field", "Value"),
            sheet_rules={
                "column_keys": (
                    "field",
                    "value",
                ),
                "workbook_name_tokens": (
                    "course_code",
                    "semester",
                    "section",
                    "academic_year",
                ),
                "required_field_keys": (
                    COURSE_METADATA_COURSE_CODE_KEY,
                    COURSE_METADATA_SEMESTER_KEY,
                    COURSE_METADATA_SECTION_KEY,
                    COURSE_METADATA_ACADEMIC_YEAR_KEY,
                    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
                ),
                "optional_field_keys": (
                    COURSE_METADATA_TOTAL_STUDENTS_KEY,
                ),
                "total_outcomes_key": COURSE_METADATA_TOTAL_OUTCOMES_KEY,
            },
        ),
        COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG: _sheet(
            key=COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
            name="Assessment_Config",
            headers=(
                "Component",
                "Weight (%)",
                "CIA",
                "CO_Wise_Marks_Breakup",
                "Direct",
                "Assessment_Type",
                "Assessment_Format",
                "Mode",
                "Participation",
            ),
            validations=[
                _rule(
                    first_row=1,
                    first_col=2,
                    last_col=2,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_VALIDATION_YES_NO_OPTIONS),
                        "input_title": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_TITLE,
                        "input_message": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_MESSAGE,
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=3,
                    last_col=3,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_VALIDATION_YES_NO_OPTIONS),
                        "input_title": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_TITLE,
                        "input_message": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_MESSAGE,
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=4,
                    last_col=4,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_VALIDATION_YES_NO_OPTIONS),
                        "input_title": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_TITLE,
                        "input_message": COURSE_SETUP_ASSESSMENT_VALIDATION_INPUT_MESSAGE,
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=5,
                    last_col=5,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS),
                        "input_title": "Assessment Type",
                        "input_message": "Select a valid Assessment_Type value",
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=6,
                    last_col=6,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS),
                        "input_title": "Assessment Format",
                        "input_message": "Select a valid Assessment_Format value",
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=7,
                    last_col=7,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_MODE_OPTIONS),
                        "input_title": "Mode",
                        "input_message": "Select a valid Mode value",
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=8,
                    last_col=8,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS),
                        "input_title": "Participation",
                        "input_message": "Select a valid Participation value",
                    },
                ),
            ],
            sheet_rules={
                "column_keys": (
                    "component",
                    "weight_percent",
                    "cia",
                    "co_wise_marks_breakup",
                    "direct",
                    "assessment_type",
                    "assessment_format",
                    "mode",
                    "participation",
                ),
                "percentage_column_keys": ("weight_percent",),
                "weight_total_expected": WEIGHT_TOTAL_EXPECTED,
                "weight_total_round_digits": WEIGHT_TOTAL_ROUND_DIGITS,
                "indirect_tools_min": 1,
                "indirect_tools_max": 3,
            },
        ),
        COURSE_SETUP_SHEET_KEY_QUESTION_MAP: _sheet(
            key=COURSE_SETUP_SHEET_KEY_QUESTION_MAP,
            name="Question_Map",
            headers=("Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO", "Bloom_Level"),
            validations=[
                _rule(
                    first_row=1,
                    first_col=4,
                    last_col=4,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS),
                        "input_title": "Bloom Level",
                        "input_message": "Select a valid Bloom_Level value",
                    },
                )
            ],
            sheet_rules={
                "column_keys": (
                    "component",
                    "question_label",
                    "max_marks",
                    "co",
                    "bloom_level",
                ),
            },
        ),
        COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION: _sheet(
            key=COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
            name="CO_Description",
            headers=("CO#", "Description", "Domain_Level", "Summary_of_Topics/Expts./Project"),
            validations=[
                _rule(
                    first_row=1,
                    first_col=0,
                    last_col=0,
                    options={
                        "validate": "integer",
                        "criteria": ">",
                        "value": 0,
                        "input_title": "CO Number",
                        "input_message": "Enter a whole number greater than zero",
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=2,
                    last_col=2,
                    options={
                        "validate": "list",
                        "source": list(COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS),
                        "input_title": "Domain Level",
                        "input_message": "Select a valid Domain_Level value",
                    },
                ),
                _rule(
                    first_row=1,
                    first_col=3,
                    last_col=3,
                    options={
                        "validate": "length",
                        "criteria": "between",
                        "minimum": COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH,
                        "maximum": COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH,
                        "input_title": "Summary Length",
                        "input_message": (
                            f"Enter text between {COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH} and "
                            f"{COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH} characters"
                        ),
                    },
                ),
            ],
        ),
        COURSE_SETUP_SHEET_KEY_STUDENTS: _sheet(
            key=COURSE_SETUP_SHEET_KEY_STUDENTS,
            name="Students",
            headers=("Reg_No", "Student_Name"),
            sheet_rules={
                "column_keys": (
                    "reg_no",
                    "student_name",
                ),
            },
        ),
    }


def _build_course_setup_blueprint(
    *,
    template_id: str,
    style_registry: dict[str, dict[str, object]],
    sheet_keys: tuple[str, ...],
) -> WorkbookBlueprint:
    catalog = _course_setup_sheet_catalog()
    missing = [key for key in sheet_keys if key not in catalog]
    if missing:
        raise ConfigurationError(f"Missing sheet definitions for keys: {', '.join(missing)}")
    sheets = [_clone_sheet(catalog[key]) for key in sheet_keys]
    return WorkbookBlueprint(
        type_id=template_id,
        style_registry={
            section: dict(values) for section, values in style_registry.items()
        },
        sheets=sheets,
        workbook_rules={
            "system_sheets": (SYSTEM_HASH_SHEET_NAME,),
            "declares_dynamic_headers": False,
            "sheet_order_enforced": True,
            "dynamic_sheet_templates": {
                CO_REPORT_SHEET_KEY_CO_INDIRECT: {
                    "sheet_name_pattern": "CO{co_index}_Indirect",
                    "header_kind": "dynamic",
                    "header_resolver": _CO_INDIRECT_HEADER_RESOLVER,
                    "header_base": ("#", "Reg. No.", "Student Name"),
                    "likert_range": (1, 5),
                    "scaled_label_template": "scaled 0-{max_value}",
                    "total_100_label": "Total (100%)",
                    "ratio_header_template": "Total ({ratio}%)",
                },
                COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE: {
                    "sheet_name_pattern": "{component_name}",
                    "header_kind": "dynamic",
                    "header_resolver": _MARKS_DIRECT_CO_WISE_HEADER_RESOLVER,
                    "header_base": MARKS_ENTRY_ROW_HEADERS,
                    "question_prefix": MARKS_ENTRY_QUESTION_PREFIX,
                    "total_label": MARKS_ENTRY_TOTAL_LABEL,
                },
                COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE: {
                    "sheet_name_pattern": "{component_name}",
                    "header_kind": "dynamic",
                    "header_resolver": _MARKS_DIRECT_NON_CO_WISE_HEADER_RESOLVER,
                    "header_base": MARKS_ENTRY_ROW_HEADERS,
                    "total_label": MARKS_ENTRY_TOTAL_LABEL,
                    "co_marks_prefix": MARKS_ENTRY_CO_MARKS_LABEL_PREFIX,
                },
                COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT: {
                    "sheet_name_pattern": "{component_name}",
                    "header_kind": "dynamic",
                    "header_resolver": _MARKS_INDIRECT_HEADER_RESOLVER,
                    "header_base": MARKS_ENTRY_ROW_HEADERS,
                    "co_prefix": CO_LABEL,
                },
            },
        },
    )


COURSE_SETUP_V1 = _build_course_setup_blueprint(
    template_id="COURSE_SETUP_V1",
    style_registry=SETUP_STYLE_REGISTRY_V1,
    sheet_keys=(
        COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
        COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
        COURSE_SETUP_SHEET_KEY_QUESTION_MAP,
        COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
        COURSE_SETUP_SHEET_KEY_STUDENTS,
    ),
)

COURSE_SETUP_V2 = _build_course_setup_blueprint(
    template_id="COURSE_SETUP_V2",
    style_registry=SETUP_STYLE_REGISTRY_V2,
    sheet_keys=(
        COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
        COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
        COURSE_SETUP_SHEET_KEY_QUESTION_MAP,
        COURSE_SETUP_SHEET_KEY_STUDENTS,
    ),
)

BLUEPRINT_REGISTRY = {
    COURSE_SETUP_V1.type_id: COURSE_SETUP_V1,
    COURSE_SETUP_V2.type_id: COURSE_SETUP_V2,
}


def get_blueprint(template_id: str) -> WorkbookBlueprint | None:
    return BLUEPRINT_REGISTRY.get(str(template_id).strip())


def get_sheet_schema(template_id: str, sheet_name: str) -> SheetSchema | None:
    blueprint = get_blueprint(template_id)
    if blueprint is None:
        return None
    target = str(sheet_name).strip()
    for sheet in blueprint.sheets:
        if sheet.name == target:
            return sheet
    return None


def get_sheet_schema_by_key(template_id: str, sheet_key: str) -> SheetSchema | None:
    blueprint = get_blueprint(template_id)
    if blueprint is None:
        return None
    target = str(sheet_key).strip()
    for sheet in blueprint.sheets:
        if sheet.key == target:
            return sheet
    # COURSE_SETUP_V2 may intentionally omit some optional sheets from its
    # default workbook blueprint while still needing schema access for other
    # workbook kinds (for example, CO description template generation).
    if str(template_id).strip() in {COURSE_SETUP_V1.type_id, COURSE_SETUP_V2.type_id}:
        catalog_sheet = _course_setup_sheet_catalog().get(target)
        if catalog_sheet is not None:
            return _clone_sheet(catalog_sheet)
    return None


def get_sheet_name_by_key(template_id: str, sheet_key: str) -> str:
    sheet = get_sheet_schema_by_key(template_id, sheet_key)
    if sheet is None:
        raise ConfigurationError(
            f"No sheet key mapping for template '{template_id}' and sheet key '{sheet_key}'."
        )
    return sheet.name


def get_sheet_headers_by_key(template_id: str, sheet_key: str) -> tuple[str, ...]:
    sheet = get_sheet_schema_by_key(template_id, sheet_key)
    if sheet is None:
        raise ConfigurationError(
            f"No sheet key mapping for template '{template_id}' and sheet key '{sheet_key}'."
        )
    if not sheet.header_matrix or not sheet.header_matrix[0]:
        raise ConfigurationError(
            f"Sheet '{sheet.name}' in template '{template_id}' has no fixed header row."
        )
    return tuple(str(value) for value in sheet.header_matrix[0])


def get_dynamic_sheet_template(template_id: str, sheet_key: str) -> dict[str, object]:
    blueprint = get_blueprint(template_id)
    if blueprint is None:
        raise ConfigurationError(f"Unknown template_id: {template_id!r}")
    dynamic_templates = blueprint.workbook_rules.get("dynamic_sheet_templates", {})
    if not isinstance(dynamic_templates, dict):
        raise ConfigurationError(
            f"Template '{template_id}' has invalid dynamic_sheet_templates declaration."
        )
    template = dynamic_templates.get(sheet_key)
    if not isinstance(template, dict):
        raise ConfigurationError(
            f"Template '{template_id}' does not define dynamic sheet template '{sheet_key}'."
        )
    return dict(template)


def _ratio_percent_token(ratio: float) -> str:
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        return f"{int(round(percent))}"
    return f"{percent:g}"


def _resolve_course_setup_co_indirect_headers(
    *,
    template_id: str,
    context: dict[str, object],
) -> tuple[str, ...]:
    dynamic_template = get_dynamic_sheet_template(template_id, CO_REPORT_SHEET_KEY_CO_INDIRECT)
    base = dynamic_template.get("header_base", ("#", "Reg. No.", "Student Name"))
    if not isinstance(base, tuple):
        raise ConfigurationError(
            f"Template '{template_id}' indirect header base must be a tuple."
        )
    likert_range = dynamic_template.get("likert_range", (1, 5))
    if (
        not isinstance(likert_range, tuple)
        or len(likert_range) != 2
        or not isinstance(likert_range[0], int)
        or not isinstance(likert_range[1], int)
    ):
        raise ConfigurationError(
            f"Template '{template_id}' indirect likert_range must be a 2-int tuple."
        )
    likert_min, likert_max = likert_range
    scaled_template = str(dynamic_template.get("scaled_label_template", "scaled 0-{max_value}"))
    total_100_label = str(dynamic_template.get("total_100_label", "Total (100%)"))
    ratio_template = str(dynamic_template.get("ratio_header_template", "Total ({ratio}%)"))
    raw_ratio = context.get("ratio", 0.2)
    if isinstance(raw_ratio, (int, float)):
        ratio = float(raw_ratio)
    elif isinstance(raw_ratio, str):
        try:
            ratio = float(raw_ratio)
        except ValueError as exc:
            raise ConfigurationError(
                f"Template '{template_id}' indirect ratio must be numeric, got {raw_ratio!r}."
            ) from exc
    else:
        raise ConfigurationError(
            f"Template '{template_id}' indirect ratio must be numeric, got {type(raw_ratio).__name__}."
        )
    ratio_token = _ratio_percent_token(ratio)

    raw_components = context.get("components", [])
    components: list[tuple[str, float]] = []
    if isinstance(raw_components, list):
        for item in raw_components:
            if isinstance(item, tuple) and len(item) == 2:
                name, weight = item
                if isinstance(name, str) and isinstance(weight, (int, float)):
                    components.append((name, float(weight)))
            elif isinstance(item, dict):
                name = item.get("name")
                weight = item.get("weight")
                if isinstance(name, str) and isinstance(weight, (int, float)):
                    components.append((name, float(weight)))
    scaled_max = max(0, likert_max - likert_min)
    has_single_component = len(components) == 1
    headers = list(base)
    for name, weight in components:
        headers.append(f"{name} ({likert_min}-{likert_max})")
        headers.append(f"{name} ({scaled_template.format(max_value=scaled_max)})")
        if not has_single_component:
            headers.append(f"{name} ({weight:g}%)")
    headers.append(total_100_label)
    headers.append(ratio_template.format(ratio=ratio_token))
    return tuple(headers)


def _positive_int_context(value: object, *, field_name: str) -> int:
    if isinstance(value, bool):
        raise ConfigurationError(f"Invalid boolean value for '{field_name}'.")
    if isinstance(value, int):
        parsed = value
    elif isinstance(value, float):
        parsed = int(value) if float(value).is_integer() else -1
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            parsed = -1
        else:
            try:
                parsed = int(text)
            except ValueError as exc:
                raise ConfigurationError(f"Invalid integer value for '{field_name}': {value!r}.") from exc
    else:
        parsed = -1
    if parsed <= 0:
        raise ConfigurationError(f"'{field_name}' must be a positive integer, got {value!r}.")
    return parsed


def _resolve_course_setup_marks_direct_co_wise_headers(
    *,
    template_id: str,
    context: dict[str, object],
) -> tuple[str, ...]:
    dynamic_template = get_dynamic_sheet_template(template_id, COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE)
    base = dynamic_template.get("header_base", MARKS_ENTRY_ROW_HEADERS)
    if not isinstance(base, tuple):
        raise ConfigurationError(f"Template '{template_id}' marks direct co-wise header base must be a tuple.")
    question_prefix = str(dynamic_template.get("question_prefix", MARKS_ENTRY_QUESTION_PREFIX))
    total_label = str(dynamic_template.get("total_label", MARKS_ENTRY_TOTAL_LABEL))
    question_count = _positive_int_context(context.get("question_count", 0), field_name="question_count")
    headers = list(base) + [f"{question_prefix}{idx}" for idx in range(1, question_count + 1)] + [total_label]
    return tuple(headers)


def _resolve_course_setup_marks_direct_non_co_wise_headers(
    *,
    template_id: str,
    context: dict[str, object],
) -> tuple[str, ...]:
    dynamic_template = get_dynamic_sheet_template(
        template_id,
        COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
    )
    base = dynamic_template.get("header_base", MARKS_ENTRY_ROW_HEADERS)
    if not isinstance(base, tuple):
        raise ConfigurationError(
            f"Template '{template_id}' marks direct non-co-wise header base must be a tuple."
        )
    total_label = str(dynamic_template.get("total_label", MARKS_ENTRY_TOTAL_LABEL))
    co_marks_prefix = str(dynamic_template.get("co_marks_prefix", MARKS_ENTRY_CO_MARKS_LABEL_PREFIX))
    raw_covered_cos = context.get("covered_cos", [])
    if not isinstance(raw_covered_cos, list):
        raise ConfigurationError(f"Template '{template_id}' marks covered_cos must be a list.")
    covered_cos: list[int] = []
    for item in raw_covered_cos:
        covered_cos.append(_positive_int_context(item, field_name="covered_cos"))
    headers = list(base) + [total_label] + [f"{co_marks_prefix}{co}" for co in covered_cos]
    return tuple(headers)


def _resolve_course_setup_marks_indirect_headers(
    *,
    template_id: str,
    context: dict[str, object],
) -> tuple[str, ...]:
    dynamic_template = get_dynamic_sheet_template(template_id, COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT)
    base = dynamic_template.get("header_base", MARKS_ENTRY_ROW_HEADERS)
    if not isinstance(base, tuple):
        raise ConfigurationError(f"Template '{template_id}' marks indirect header base must be a tuple.")
    co_prefix = str(dynamic_template.get("co_prefix", CO_LABEL))
    total_outcomes = _positive_int_context(context.get("total_outcomes", 0), field_name="total_outcomes")
    headers = list(base) + [f"{co_prefix}{i}" for i in range(1, total_outcomes + 1)]
    return tuple(headers)


def resolve_dynamic_sheet_headers(
    template_id: str,
    *,
    sheet_key: str,
    context: dict[str, object] | None = None,
) -> tuple[str, ...]:
    dynamic_template = get_dynamic_sheet_template(template_id, sheet_key)
    resolver_name = str(dynamic_template.get("header_resolver", "")).strip()
    resolved_context = dict(context or {})
    if resolver_name == _CO_INDIRECT_HEADER_RESOLVER and sheet_key == CO_REPORT_SHEET_KEY_CO_INDIRECT:
        return _resolve_course_setup_co_indirect_headers(
            template_id=template_id,
            context=resolved_context,
        )
    if (
        resolver_name == _MARKS_DIRECT_CO_WISE_HEADER_RESOLVER
        and sheet_key == COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE
    ):
        return _resolve_course_setup_marks_direct_co_wise_headers(
            template_id=template_id,
            context=resolved_context,
        )
    if (
        resolver_name == _MARKS_DIRECT_NON_CO_WISE_HEADER_RESOLVER
        and sheet_key == COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE
    ):
        return _resolve_course_setup_marks_direct_non_co_wise_headers(
            template_id=template_id,
            context=resolved_context,
        )
    if resolver_name == _MARKS_INDIRECT_HEADER_RESOLVER and sheet_key == COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT:
        return _resolve_course_setup_marks_indirect_headers(
            template_id=template_id,
            context=resolved_context,
        )
    raise ConfigurationError(
        f"Template '{template_id}' dynamic header resolver not supported for sheet '{sheet_key}': {resolver_name!r}"
    )


def get_system_hash_sheet_schema() -> dict[str, object]:
    return {
        "name": SYSTEM_HASH_SHEET_NAME,
        "headers": (
            SYSTEM_HASH_HEADER_TEMPLATE_ID,
            SYSTEM_HASH_HEADER_TEMPLATE_HASH,
        ),
        "keys": (
            SYSTEM_HASH_KEY_TEMPLATE_ID,
            SYSTEM_HASH_KEY_TEMPLATE_HASH,
        ),
    }
