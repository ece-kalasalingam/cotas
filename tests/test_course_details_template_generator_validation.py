from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import (
    ASSESSMENT_CONFIG_SHEET,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    ID_COURSE_SETUP,
    QUESTION_MAP_SHEET,
    STUDENTS_SHEET,
    SYSTEM_HASH_SHEET,
)
from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.texts import get_language, set_language
from domain import instructor_template_engine as mod
from domain.instructor_template_engine import (
    generate_course_details_template,
    validate_course_details_workbook,
)


def _build_workbook(tmp_path: Path) -> Path:
    output = tmp_path / "course_setup.xlsx"
    generate_course_details_template(output)
    return output


def test_validate_uploaded_workbook_accepts_generated_template(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    assert validate_course_details_workbook(output) == ID_COURSE_SETUP


def test_validate_uploaded_workbook_rejects_hash_tampering(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["__SYSTEM_HASH__"]["B2"] = "bad-hash"
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="Template hash mismatch"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_bad_direct_weight_total(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Assessment_Config"]["B2"] = 1
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="direct component weights must total 100"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_multi_co_for_co_wise_component(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Question_Map"]["D2"] = "CO1, CO2"
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="requires exactly one CO per question"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_duplicate_student_reg_no(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Students"]["A3"] = workbook["Students"]["A2"].value
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="duplicate Reg_No"):
        validate_course_details_workbook(output)


def test_generation_is_schema_stable_when_language_changes_mid_process(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    output = tmp_path / "course_setup.xlsx"
    original_generate_worksheet = mod.generate_worksheet
    switched = {"done": False}
    previous_language = get_language()

    def _generate_worksheet_with_mid_switch(*args, **kwargs):
        if not switched["done"]:
            switched["done"] = True
            set_language("ta-IN")
        return original_generate_worksheet(*args, **kwargs)

    set_language("en")
    monkeypatch.setattr(mod, "generate_worksheet", _generate_worksheet_with_mid_switch)
    try:
        generate_course_details_template(output)
        assert switched["done"] is True
        assert validate_course_details_workbook(output) == ID_COURSE_SETUP

        workbook = openpyxl.load_workbook(output, data_only=True)
        try:
            assert workbook.sheetnames == [
                COURSE_METADATA_SHEET,
                ASSESSMENT_CONFIG_SHEET,
                QUESTION_MAP_SHEET,
                STUDENTS_SHEET,
                SYSTEM_HASH_SHEET,
            ]
            assert workbook[COURSE_METADATA_SHEET]["A1"].value == COURSE_METADATA_HEADERS[0]
            assert workbook[COURSE_METADATA_SHEET]["B1"].value == COURSE_METADATA_HEADERS[1]
        finally:
            workbook.close()
    finally:
        set_language(previous_language)


def test_generate_course_details_template_honors_pre_cancel(tmp_path: Path) -> None:
    output = tmp_path / "cancelled_course_setup.xlsx"
    token = CancellationToken()
    token.cancel()

    with pytest.raises(JobCancelledError):
        generate_course_details_template(output, cancel_token=token)
