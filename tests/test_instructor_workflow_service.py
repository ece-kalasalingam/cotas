from __future__ import annotations

import logging
import time
from pathlib import Path
from threading import Event

import pytest

from common.exceptions import AppSystemError, JobCancelledError
from common.jobs import CancellationToken
from services import instructor_workflow_service as service_mod


def test_service_honors_pre_cancel_before_generation(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="generate_course_template")

    called = {"count": 0}

    def _fake_generate(**_kwargs):
        called["count"] += 1
        return Path("unused.xlsx")

    monkeypatch.setattr(service_mod, "generate_workbook", _fake_generate)

    with pytest.raises(JobCancelledError):
        service.generate_course_details_template(tmp_path / "out.xlsx", context=context, cancel_token=token)

    assert called["count"] == 0


def test_service_generate_final_report_copies_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")

    called: dict[str, object] = {}

    def _fake_resolve(_self: object, _path: str | Path) -> str:
        return "COURSE_SETUP_V2"

    def _fake_generate_workbook(**kwargs) -> Path:
        called["source"] = kwargs.get("context", {}).get("filled_marks_path")
        called["output"] = kwargs.get("output_path")
        output = Path(kwargs["output_path"])
        output.write_text("generated", encoding="utf-8")
        return output

    monkeypatch.setattr(service_mod.InstructorWorkflowService, "_resolve_template_id_from_workbook", _fake_resolve)
    monkeypatch.setattr(service_mod, "generate_workbook", _fake_generate_workbook)

    result = service.generate_final_report(src, dst, context=context)

    assert result == dst
    assert called["source"] == str(src)
    assert called["output"] == dst
    assert dst.read_text(encoding="utf-8") == "generated"


def test_service_logs_step_lifecycle(
    caplog: pytest.LogCaptureFixture, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")
    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )
    monkeypatch.setattr(service_mod, "generate_workbook", lambda **_kwargs: dst)

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    service.generate_final_report(src, dst, context=context)

    messages = [record.getMessage() for record in caplog.records]
    assert "Instructor workflow step started." in messages
    assert "Instructor workflow step completed." in messages
    step_records = [
        record
        for record in caplog.records
        if getattr(record, "step_id", None) == "generate_final_report"
    ]
    assert step_records


def test_service_logs_cancellation(caplog: pytest.LogCaptureFixture, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="generate_course_template")

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(JobCancelledError):
        service.generate_course_details_template(tmp_path / "out.xlsx", context=context, cancel_token=token)

    assert any(record.getMessage() == "Instructor workflow step cancelled." for record in caplog.records)


def test_service_generate_final_report_keeps_existing_dest_on_copy_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("fresh", encoding="utf-8")
    dst.write_text("stable", encoding="utf-8")

    def _failing_generate(**_kwargs) -> Path:
        raise OSError("generation interrupted")

    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )
    monkeypatch.setattr(service_mod, "generate_workbook", _failing_generate)

    with pytest.raises(OSError):
        service.generate_final_report(src, dst, context=context)

    assert dst.read_text(encoding="utf-8") == "stable"


def test_service_enforces_timeout(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")
    monkeypatch.setenv("FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS", "1")

    def _slow_generate(**_kwargs):
        time.sleep(2)
        return dst

    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )
    monkeypatch.setattr(service_mod, "generate_workbook", _slow_generate)

    with pytest.raises(AppSystemError, match="exceeded timeout"):
        service.generate_final_report(src, dst, context=context)


def test_service_timeout_prevents_post_timeout_output_mutation(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")
    dst.write_text("stable", encoding="utf-8")
    monkeypatch.setenv("FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS", "1")

    state = {"token_seen": False}
    worker_done = Event()

    def _slow_generate_with_cancel(**kwargs) -> Path:
        state["token_seen"] = kwargs.get("cancel_token") is not None
        try:
            time.sleep(1.2)
            cancel_token = kwargs.get("cancel_token")
            if isinstance(cancel_token, CancellationToken):
                cancel_token.raise_if_cancelled()
            output = Path(kwargs["output_path"])
            output.write_text("late-write", encoding="utf-8")
            return output
        finally:
            worker_done.set()

    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )
    monkeypatch.setattr(service_mod, "generate_workbook", _slow_generate_with_cancel)

    with pytest.raises(AppSystemError, match="exceeded timeout"):
        service.generate_final_report(src, dst, context=context)

    assert worker_done.wait(timeout=2.5)
    assert state["token_seen"] is True
    assert dst.read_text(encoding="utf-8") == "stable"


def test_service_logs_stable_error_code_for_validation_errors(
    caplog: pytest.LogCaptureFixture, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="validate_course_template")
    monkeypatch.setattr(
        service_mod,
        "validate_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            service_mod.ValidationError("bad data", code="BAD_DATA")
        ),
    )

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(service_mod.ValidationError):
        service.validate_course_details_workbook(tmp_path / "broken.xlsx", context=context)

    codes = [getattr(record, "error_code", None) for record in caplog.records]
    assert "BAD_DATA" in codes


def test_resolve_timeout_seconds_invalid_env_uses_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, "not-an-int")
    assert (
        service_mod.InstructorWorkflowService._resolve_timeout_seconds()
        == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS
    )


def test_call_with_optional_cancel_token_signature_fallback_branch(monkeypatch: pytest.MonkeyPatch) -> None:
    # Force `signature(fn)` to fail so fallback Signature() path is exercised.
    monkeypatch.setattr(
        service_mod,
        "signature",
        lambda _fn: (_ for _ in ()).throw(ValueError("bad signature")),
    )
    called: list[tuple[object, ...]] = []

    def _fn_no_cancel(*args):
        called.append(args)
        return "ok"

    out = service_mod.InstructorWorkflowService._call_with_optional_cancel_token(
        _fn_no_cancel,
        "a",
        "b",
        cancel_token=CancellationToken(),
    )

    assert out == "ok"
    assert called == [("a", "b")]

