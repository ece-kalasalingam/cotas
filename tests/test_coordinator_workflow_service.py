from __future__ import annotations

import time
from pathlib import Path

import pytest

from common.exceptions import AppSystemError, JobCancelledError
from common.jobs import CancellationToken
from services import coordinator_workflow_service as service_mod


def test_resolve_timeout_seconds_defaults_for_missing_invalid_and_non_positive(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, raising=False)
    assert service_mod.CoordinatorWorkflowService._resolve_timeout_seconds() == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS

    monkeypatch.setenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, "abc")
    assert service_mod.CoordinatorWorkflowService._resolve_timeout_seconds() == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS

    monkeypatch.setenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, "0")
    assert service_mod.CoordinatorWorkflowService._resolve_timeout_seconds() == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS


def test_collect_files_honors_pre_cancel_before_work() -> None:
    service = service_mod.CoordinatorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="collect")

    called = {"count": 0}

    def _fake_analyze(*_args, **_kwargs):
        called["count"] += 1
        return {}
    monkeypatch = pytest.MonkeyPatch()
    monkeypatch.setattr(service_mod, "_analyze_dropped_files", _fake_analyze)

    try:
        with pytest.raises(JobCancelledError):
            service.collect_files(
                ["a.xlsx"],
                existing_keys=set(),
                existing_paths=[],
                context=context,
                cancel_token=token,
            )
    finally:
        monkeypatch.undo()

    assert called["count"] == 0


def test_calculate_attainment_passes_token_to_generator(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    service = service_mod.CoordinatorWorkflowService()
    context = service.create_job_context(step_id="calc")
    source = [tmp_path / "in.xlsx"]
    out = tmp_path / "out.xlsx"

    seen = {"token": None}

    def _fake_generate(
        src_paths,
        output_path,
        *,
        token=None,
        thresholds=None,
        co_attainment_percent=None,
        co_attainment_level=None,
    ):
        seen["token"] = token
        return output_path
    monkeypatch.setattr(service_mod, "_generate_co_attainment_workbook", _fake_generate)

    result = service.calculate_attainment(
        source,
        out,
        context=context,
        cancel_token=CancellationToken(),
    )

    assert result == out
    assert seen["token"] is not None


def test_run_with_timeout_cancels_token_and_raises_system_error() -> None:
    token = CancellationToken()

    with pytest.raises(AppSystemError, match="exceeded timeout"):
        service_mod.CoordinatorWorkflowService._run_with_timeout(
            operation="op",
            work=lambda: time.sleep(1.2),
            timeout_seconds=1,
            cancel_token=token,
        )

    with pytest.raises(JobCancelledError):
        token.raise_if_cancelled()


def test_execute_with_telemetry_app_system_and_unexpected_error_paths() -> None:
    service = service_mod.CoordinatorWorkflowService()
    context = service.create_job_context(step_id="collect")

    monkeypatch = pytest.MonkeyPatch()
    try:
        monkeypatch.setattr(
            service_mod,
            "_analyze_dropped_files",
            lambda *_a, **_k: (_ for _ in ()).throw(AppSystemError("system")),
        )
        with pytest.raises(AppSystemError):
            service.collect_files(
                ["a.xlsx"],
                existing_keys=set(),
                existing_paths=[],
                context=context,
                cancel_token=CancellationToken(),
            )

        monkeypatch.setattr(
            service_mod,
            "_analyze_dropped_files",
            lambda *_a, **_k: (_ for _ in ()).throw(ValueError("boom")),
        )
        with pytest.raises(ValueError):
            service.collect_files(
                ["a.xlsx"],
                existing_keys=set(),
                existing_paths=[],
                context=context,
                cancel_token=CancellationToken(),
            )
    finally:
        monkeypatch.undo()
