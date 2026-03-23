"""COURSE_SETUP_V2 marks-template generation."""

from __future__ import annotations

from typing import Any

from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.registry import (
    COURSE_METADATA_TOTAL_STUDENTS_KEY,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    COURSE_SETUP_SHEET_KEY_STUDENTS,
    get_blueprint,
    get_sheet_name_by_key,
    get_sheet_schema_by_key,
)
from common.utils import coerce_excel_number, normalize
from domain.template_versions.course_setup_v2_impl import instructor_engine_shareops as _shareops
#
# _add_system_layout_sheet = _sheetops._add_system_layout_sheet
# _copy_system_hash_sheet = _sheetops._copy_system_hash_sheet
_ensure_uniform_template_id_and_copy_system_hash = _shareops.ensure_uniform_template_id_and_copy_system_hash


def ensure_uniform_template_id_and_copy_system_hash(
    source_workbooks: list[Any],
    target_workbook: Any,
    *,
    routed_template_id: str | None = None,
    cancel_token: CancellationToken | None = None,
) -> str:
    return _ensure_uniform_template_id_and_copy_system_hash(
        source_workbooks,
        target_workbook,
        routed_template_id=routed_template_id,
        cancel_token=cancel_token,
    )


def generate_students_sheet_from_sections(
    source_workbooks: list[Any],
    target_workbook: Any,
    *,
    template_id: str,
    cancel_token: CancellationToken | None = None,
) -> list[tuple[str, str]]:
    schema = get_sheet_schema_by_key(template_id, COURSE_SETUP_SHEET_KEY_STUDENTS)
    if schema is None or not schema.header_matrix or not schema.header_matrix[0]:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_STUDENTS),
        )
    headers = list(schema.header_matrix[0])
    if len(headers) < 2:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_COLUMN_KEY_MISSING",
            sheet_name=schema.name,
            column_key="reg_no/student_name",
        )

    seen_reg: set[str] = set()
    students: list[tuple[str, str]] = []
    for workbook in source_workbooks:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        if schema.name not in getattr(workbook, "sheetnames", []):
            raise validation_error_from_key(
                "validation.system.sheet_missing",
                code="COA_SYSTEM_SHEET_MISSING",
                sheet=schema.name,
            )
        worksheet = workbook[schema.name]
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=2, max_col=2, values_only=True),
            start=2,
        ):
            reg_raw = row[0] if len(row) > 0 else None
            name_raw = row[1] if len(row) > 1 else None
            reg_no = str(reg_raw).strip() if reg_raw is not None else ""
            student_name = str(name_raw).strip() if name_raw is not None else ""
            if not reg_no and not student_name:
                continue
            if not reg_no or not student_name:
                raise validation_error_from_key(
                    "instructor.validation.students_reg_and_name_required",
                    row=row_number,
                )
            reg_key = normalize(reg_no)
            if reg_key in seen_reg:
                raise validation_error_from_key(
                    "instructor.validation.students_duplicate_reg_no",
                    reg_no=reg_no,
                    code="STUDENTS_DUPLICATE_REG_NO",
                )
            seen_reg.add(reg_key)
            students.append((reg_no, student_name))

    blueprint = get_blueprint(template_id)
    if blueprint is None:
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
        )
    header_format = _shareops.build_header_format(target_workbook, blueprint.style_registry.get("header", {}))
    body_format = _shareops.build_body_format(target_workbook, blueprint.style_registry.get("body", {}))
    _shareops.generate_worksheet(
        workbook=target_workbook,
        sheet_name=schema.name,
        headers=headers,
        data=students,
        header_format=header_format,
        body_format=body_format,
    )
    return students


def generate_course_metadata_sheet_from_sections(
    source_workbooks: list[Any],
    target_workbook: Any,
    *,
    template_id: str,
    total_students: int,
    cancel_token: CancellationToken | None = None,
) -> list[tuple[str, str]]:
    schema = get_sheet_schema_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    if schema is None or not schema.header_matrix or not schema.header_matrix[0]:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA),
        )
    headers = list(schema.header_matrix[0])
    required_keys_raw = schema.sheet_rules.get("required_field_keys")
    required_keys = [
        normalize(item)
        for item in (required_keys_raw if isinstance(required_keys_raw, (list, tuple)) else [])
        if isinstance(item, str) and normalize(item)
    ]
    if not required_keys:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_COLUMN_KEY_MISSING",
            sheet_name=schema.name,
            column_key="required_field_keys",
        )
    optional_keys_raw = schema.sheet_rules.get("optional_field_keys")
    optional_keys = [
        normalize(item)
        for item in (optional_keys_raw if isinstance(optional_keys_raw, (list, tuple)) else [])
        if isinstance(item, str) and normalize(item)
    ]

    merged_values: dict[str, list[str]] = {key: [] for key in required_keys}
    for workbook in source_workbooks:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        if schema.name not in getattr(workbook, "sheetnames", []):
            raise validation_error_from_key(
                "validation.system.sheet_missing",
                code="COA_SYSTEM_SHEET_MISSING",
                sheet=schema.name,
            )
        worksheet = workbook[schema.name]
        current: dict[str, str] = {}
        for row in worksheet.iter_rows(min_row=2, max_col=2, values_only=True):
            key_raw = row[0] if len(row) > 0 else None
            value_raw = row[1] if len(row) > 1 else None
            key = normalize(key_raw)
            if not key:
                continue
            value = coerce_excel_number(value_raw)
            current[key] = str(value).strip() if value is not None else ""
        missing = [key for key in required_keys if not current.get(key, "").strip()]
        if missing:
            raise validation_error_from_key(
                "instructor.validation.course_metadata_missing_fields",
                fields=", ".join(missing),
            )
        for key in required_keys:
            value = current[key]
            bucket = merged_values[key]
            if value not in bucket:
                bucket.append(value)

    rows: list[tuple[str, str]] = []
    for key in required_keys:
        values = merged_values[key]
        rows.append((key, values[0] if len(values) == 1 else ", ".join(values)))

    if COURSE_METADATA_TOTAL_STUDENTS_KEY in optional_keys:
        rows.append((COURSE_METADATA_TOTAL_STUDENTS_KEY, str(int(total_students))))

    blueprint = get_blueprint(template_id)
    if blueprint is None:
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
        )
    header_format = _shareops.build_header_format(target_workbook, blueprint.style_registry.get("header", {}))
    body_format = _shareops.build_body_format(target_workbook, blueprint.style_registry.get("body", {}))
    _shareops.generate_worksheet(
        workbook=target_workbook,
        sheet_name=schema.name,
        headers=headers,
        data=rows,
        header_format=header_format,
        body_format=body_format,
    )
    return rows
#
#
# def _ve(translation_key: str, *, code: str, **context: Any) -> ValidationError:
#     return validation_error_from_key(translation_key, code=code, **context)
#
#
# def generate_marks_template_from_course_details(
#     course_details_path: str | Path,
#     output_path: str | Path,
#     *,
#     cancel_token: CancellationToken | None = None,
# ) -> Path:
#     """Generate marks-entry workbook from a validated course-details workbook."""
#     try:
#         import openpyxl
#     except ModuleNotFoundError as exc:
#         raise _ve(
#             "instructor.validation.openpyxl_missing",
#             code="OPENPYXL_MISSING",
#         ) from exc
#
#     try:
#         import xlsxwriter
#     except ModuleNotFoundError as exc:
#         raise _ve(
#             "instructor.validation.xlsxwriter_missing",
#             code="XLSXWRITER_MISSING",
#         ) from exc
#
#     source_file = Path(course_details_path)
#     if not source_file.exists():
#         raise _ve(
#             "instructor.validation.workbook_not_found",
#             code="WORKBOOK_NOT_FOUND",
#             workbook=str(source_file),
#         )
#
#     output = Path(output_path)
#     output.parent.mkdir(parents=True, exist_ok=True)
#     temp_output = output.with_name(f"{output.name}.{uuid4().hex}{WORKBOOK_TEMP_SUFFIX}")
#
#     try:
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         workbook = openpyxl.load_workbook(source_file, data_only=True)
#     except JobCancelledError:
#         raise
#     except Exception as exc:
#         raise _ve(
#             "instructor.validation.workbook_open_failed",
#             code="WORKBOOK_OPEN_FAILED",
#             workbook=str(source_file),
#         ) from exc
#
#     target = xlsxwriter.Workbook(str(temp_output), {"constant_memory": True})
#     target_closed = False
#
#     def _cleanup_incomplete_output() -> None:
#         nonlocal target_closed
#         if not target_closed:
#             try:
#                 target.close()
#                 target_closed = True
#             except Exception:
#                 _logger.debug("Suppressing target close error during cleanup.", exc_info=True)
#         if temp_output.exists():
#             try:
#                 temp_output.unlink()
#             except OSError:
#                 _logger.warning("Failed to cleanup temp marks template file: %s", temp_output)
#
#     try:
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         template_id = _extract_and_validate_template_id(workbook)
#         context = _extract_marks_template_context_by_template(workbook, template_id)
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         layout_manifest = _write_marks_template_workbook_by_template(
#             target,
#             context,
#             template_id=template_id,
#             cancel_token=cancel_token,
#         )
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         _copy_system_hash_sheet(workbook, target)
#         _add_system_layout_sheet(target, layout_manifest)
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         target.close()
#         target_closed = True
#         if cancel_token is not None:
#             cancel_token.raise_if_cancelled()
#         os.replace(temp_output, output)
#     except ValidationError:
#         _cleanup_incomplete_output()
#         raise
#     except JobCancelledError:
#         _cleanup_incomplete_output()
#         raise
#     except Exception as exc:
#         _cleanup_incomplete_output()
#         _logger.exception(
#             "Failed to generate marks template. source=%s output=%s",
#             source_file,
#             output,
#         )
#         raise AppSystemError(
#             t("instructor.system.template_generate_failed", output=output)
#         ) from exc
#     finally:
#         workbook.close()
#
#     return output
#
#
# def _extract_and_validate_template_id(workbook: Any) -> str:
#     return read_valid_template_id_from_system_hash_sheet(workbook)
#
#
# def _extract_marks_template_context_by_template(workbook: Any, template_id: str) -> dict[str, Any]:
#     return extract_marks_template_context_by_template(workbook, template_id=template_id)
#
#
# def _write_marks_template_workbook_by_template(
#     workbook: Any,
#     context: dict[str, Any],
#     *,
#     template_id: str,
#     cancel_token: CancellationToken | None = None,
# ) -> dict[str, Any]:
#     return write_marks_template_workbook_by_template(
#         workbook,
#         context,
#         template_id=template_id,
#         cancel_token=cancel_token,
#     )
#
#
# __all__ = ["generate_marks_template_from_course_details"]

# Placeholder exports while existing implementation stays fully commented.
__all__ = [
    "ensure_uniform_template_id_and_copy_system_hash",
    "generate_students_sheet_from_sections",
    "generate_course_metadata_sheet_from_sections",
]
