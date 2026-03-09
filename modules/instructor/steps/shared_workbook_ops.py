"""Shared workbook operations for instructor workflow steps."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
from pathlib import Path

from common.constants import COURSE_METADATA_SHEET
from common.texts import t
from common.utils import coerce_excel_number, normalize


def sanitize_filename_token(value: object) -> str:
    token = str(value).strip()
    token = re.sub(r'[<>:"/\\|?*]+', "_", token)
    token = re.sub(r"\s+", "", token)
    token = token.strip(" ._")
    return token


def build_marks_template_default_name(course_details_path: str | None) -> str:
    fallback = t("instructor.dialog.step3.default_name")
    if not course_details_path:
        return fallback

    try:
        import openpyxl
    except ModuleNotFoundError:
        return fallback

    workbook = None
    try:
        workbook = openpyxl.load_workbook(course_details_path, data_only=True)
        if COURSE_METADATA_SHEET not in workbook.sheetnames:
            return fallback
        sheet = workbook[COURSE_METADATA_SHEET]
        fields: dict[str, str] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            key = normalize(row[0] if len(row) > 0 else None)
            if not key:
                continue
            value = row[1] if len(row) > 1 else None
            coerced = coerce_excel_number(value)
            fields[key] = str(coerced).strip() if coerced is not None else ""

        parts = [
            sanitize_filename_token(fields.get("course_code", "")),
            sanitize_filename_token(fields.get("semester", "")),
            sanitize_filename_token(fields.get("section", "")),
            sanitize_filename_token(fields.get("academic_year", "")),
            "Marks",
        ]
        if any(not part for part in parts[:4]):
            return fallback
        return f"{'_'.join(parts)}.xlsx"
    except Exception:
        return fallback
    finally:
        if workbook is not None:
            workbook.close()


def atomic_copy_file(source_path: str | Path, output_path: str | Path, *, logger: object | None = None) -> Path:
    source = Path(source_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_name = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            delete=False,
            dir=str(output.parent),
            prefix=f"{output.name}.",
            suffix=".tmp",
        ) as temp_file:
            temp_name = temp_file.name
        shutil.copyfile(str(source), temp_name)
        os.replace(temp_name, output)
    except Exception:
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                if logger is not None:
                    warning = getattr(logger, "warning", None)
                    if callable(warning):
                        warning("Failed to cleanup temp report file: %s", temp_name)
        raise
    return output
