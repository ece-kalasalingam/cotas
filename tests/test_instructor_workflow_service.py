from __future__ import annotations

import logging
import time
from pathlib import Path

import pytest

from common.exceptions import AppSystemError, JobCancelledError
from common.jobs import CancellationToken
from services import instructor_workflow_service as service_mod


def test_service_honors_pre_cancel_before_generation(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="step1")

    called = {"count": 0}

    def _fake_generate(_path):
        called["count"] += 1
        return Path("unused.xlsx")

    monkeypatch.setattr(service_mod, "generate_course_details_template", _fake_generate)

    with pytest.raises(JobCancelledError):
        service.generate_course_details_template(tmp_path / "out.xlsx", context=context, cancel_token=token)

    assert called["count"] == 0


def test_service_generate_final_report_copies_file(tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step3")
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")

    result = service.generate_final_report(src, dst, context=context)

    assert result == dst
    assert dst.read_text(encoding="utf-8") == "data"


def test_service_logs_step_lifecycle(caplog: pytest.LogCaptureFixture, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step3")
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    service.generate_final_report(src, dst, context=context)

    messages = [record.getMessage() for record in caplog.records]
    assert "Instructor workflow step started." in messages
    assert "Instructor workflow step completed." in messages
    step_records = [record for record in caplog.records if getattr(record, "step_id", None) == "step3"]
    assert step_records


def test_service_logs_cancellation(caplog: pytest.LogCaptureFixture, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="step1")

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(JobCancelledError):
        service.generate_course_details_template(tmp_path / "out.xlsx", context=context, cancel_token=token)

    assert any(record.getMessage() == "Instructor workflow step cancelled." for record in caplog.records)


def test_service_generate_final_report_keeps_existing_dest_on_copy_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step3")
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("fresh", encoding="utf-8")
    dst.write_text("stable", encoding="utf-8")

    def _failing_copy(_src: str, dst_path: str) -> None:
        Path(dst_path).write_text("partial", encoding="utf-8")
        raise OSError("copy interrupted")

    monkeypatch.setattr(service_mod.shutil, "copyfile", _failing_copy)

    with pytest.raises(OSError):
        service.generate_final_report(src, dst, context=context)

    assert dst.read_text(encoding="utf-8") == "stable"
    assert list(tmp_path.glob("report.xlsx.*.tmp")) == []


def test_service_generate_final_report_cleans_temp_when_replace_fails(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step3")
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("fresh", encoding="utf-8")
    dst.write_text("stable", encoding="utf-8")

    monkeypatch.setattr(service_mod.os, "replace", lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("replace failed")))

    with pytest.raises(OSError):
        service.generate_final_report(src, dst, context=context)

    assert dst.read_text(encoding="utf-8") == "stable"
    assert list(tmp_path.glob("report.xlsx.*.tmp")) == []


def test_service_enforces_timeout(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step3")
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")
    monkeypatch.setenv("FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS", "1")

    def _slow_copy(*_args, **_kwargs):
        time.sleep(2)
        return None

    monkeypatch.setattr(service_mod.shutil, "copyfile", _slow_copy)

    with pytest.raises(AppSystemError, match="exceeded timeout"):
        service.generate_final_report(src, dst, context=context)


def test_service_logs_stable_error_code_for_validation_errors(
    caplog: pytest.LogCaptureFixture, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="step2")
    monkeypatch.setattr(
        service_mod,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            service_mod.ValidationError("bad data", code="BAD_DATA")
        ),
    )

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(service_mod.ValidationError):
        service.validate_course_details_workbook(tmp_path / "broken.xlsx", context=context)

    codes = [getattr(record, "error_code", None) for record in caplog.records]
    assert "BAD_DATA" in codes
