from __future__ import annotations

import logging
import time
from pathlib import Path
from threading import Event

import pytest

from common.exceptions import AppSystemError, JobCancelledError
from common.jobs import CancellationToken
from services import instructor_workflow_service as service_mod


def test_service_generate_final_report_copies_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Test service generate final report copies file.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")

    called: dict[str, object] = {}

    def _fake_resolve(_self: object, _path: str | Path) -> str:
        """Fake resolve.
        
        Args:
            _self: Parameter value (object).
            _path: Parameter value (str | Path).
        
        Returns:
            str: Return value.
        
        Raises:
            None.
        """
        return "COURSE_SETUP_V2"

    def _fake_generate_workbook(**kwargs) -> Path:
        """Fake generate workbook.
        
        Args:
            kwargs: Parameter value.
        
        Returns:
            Path: Return value.
        
        Raises:
            None.
        """
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
    """Test service logs step lifecycle.
    
    Args:
        caplog: Parameter value (pytest.LogCaptureFixture).
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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


def test_service_logs_cancellation(
    caplog: pytest.LogCaptureFixture,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test service logs cancellation.
    
    Args:
        caplog: Parameter value (pytest.LogCaptureFixture).
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="prepare_marks")
    source = tmp_path / "course_details.xlsx"
    output = tmp_path / "marks_template.xlsx"
    called = {"count": 0}

    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )

    def _fake_generate_workbooks(**_kwargs):
        """Fake generate workbooks.
        
        Args:
            _kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        called["count"] += 1
        return {"generated_workbook_paths": []}

    monkeypatch.setattr(service_mod, "generate_workbooks", _fake_generate_workbooks)

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(JobCancelledError):
        service.generate_marks_template(source, output, context=context, cancel_token=token)

    assert any(record.getMessage() == "Instructor workflow step cancelled." for record in caplog.records)
    assert called["count"] == 0


def test_service_generate_final_report_keeps_existing_dest_on_copy_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Test service generate final report keeps existing dest on copy failure.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("fresh", encoding="utf-8")
    dst.write_text("stable", encoding="utf-8")

    def _failing_generate(**_kwargs) -> Path:
        """Failing generate.
        
        Args:
            _kwargs: Parameter value.
        
        Returns:
            Path: Return value.
        
        Raises:
            None.
        """
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
    """Test service enforces timeout.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    src = tmp_path / "filled.txt"
    dst = tmp_path / "report.xlsx"
    src.write_text("data", encoding="utf-8")
    monkeypatch.setenv("FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS", "1")

    def _slow_generate(**_kwargs):
        """Slow generate.
        
        Args:
            _kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
    """Test service timeout prevents post timeout output mutation.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
        """Slow generate with cancel.
        
        Args:
            kwargs: Parameter value.
        
        Returns:
            Path: Return value.
        
        Raises:
            None.
        """
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
    """Test service logs stable error code for validation errors.
    
    Args:
        caplog: Parameter value (pytest.LogCaptureFixture).
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    context = service.create_job_context(step_id="generate_final_report")
    monkeypatch.setattr(
        service_mod,
        "generate_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            service_mod.ValidationError("bad data", code="BAD_DATA")
        ),
    )

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(service_mod.ValidationError):
        src = tmp_path / "filled.txt"
        dst = tmp_path / "report.xlsx"
        src.write_text("data", encoding="utf-8")
        monkeypatch.setattr(
            service_mod.InstructorWorkflowService,
            "_resolve_template_id_from_workbook",
            lambda _self, _path: "COURSE_SETUP_V2",
        )
        service.generate_final_report(src, dst, context=context)

    codes = [getattr(record, "error_code", None) for record in caplog.records]
    assert "BAD_DATA" in codes


def test_resolve_timeout_seconds_invalid_env_uses_default(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test resolve timeout seconds invalid env uses default.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, "not-an-int")
    assert (
        service_mod.InstructorWorkflowService._resolve_timeout_seconds()
        == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS
    )


def test_call_with_optional_cancel_token_signature_fallback_branch(monkeypatch: pytest.MonkeyPatch) -> None:
    # Force `signature(fn)` to fail so fallback Signature() path is exercised.
    """Test call with optional cancel token signature fallback branch.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(
        service_mod,
        "signature",
        lambda _fn: (_ for _ in ()).throw(ValueError("bad signature")),
    )
    called: list[tuple[object, ...]] = []

    def _fn_no_cancel(*args):
        """Fn no cancel.
        
        Args:
            args: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
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



