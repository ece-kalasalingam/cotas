from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, cast

from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from modules.instructor.steps import step2_filled_marks_and_final_report as step2_phase


class _Token:
    def __init__(self) -> None:
        self.cancelled = False

    def raise_if_cancelled(self) -> None:
        if self.cancelled:
            raise JobCancelledError("cancelled")


@dataclass
class _Ctx:
    job_id: str
    step_id: str


class _Logger:
    def __init__(self) -> None:
        self.info_calls: list[tuple] = []
        self.warning_calls: list[tuple] = []

    def info(self, *args, **kwargs) -> None:
        self.info_calls.append((args, kwargs))

    def warning(self, *args, **kwargs) -> None:
        self.warning_calls.append((args, kwargs))


class _WorkflowService:
    def __init__(self) -> None:
        self.generated: list[tuple[str, str]] = []

    def create_job_context(self, *, step_id: str, payload: dict[str, object]) -> _Ctx:
        del payload
        return _Ctx(job_id="job-s3", step_id=step_id)

    def generate_final_report(
        self,
        source_path: str,
        save_path: str,
        *,
        context: _Ctx,
        cancel_token: _Token,
    ) -> Path:
        del context, cancel_token
        self.generated.append((source_path, save_path))
        return Path(save_path)


class _Module:
    @dataclass
    class _State:
        busy: bool = False

    def __init__(self) -> None:
        self.state = _Module._State()
        self.filled_marks_done = False
        self.filled_marks_path: str | None = None
        self.filled_marks_outdated = False
        self.final_report_outdated = False
        self.final_report_path: str | None = None
        self.final_report_done = False
        self._workflow_service: _WorkflowService | None = None
        self.can_run = True
        self.can_run_reason = "blocked"
        self.published: list[str] = []
        self.toasts: list[tuple[str, str]] = []
        self.remembered: list[str] = []
        self.started: dict[str, object] = {}

    def _remember_dialog_dir_safe(self, path: str) -> None:
        self.remembered.append(path)

    def _show_validation_error_toast(self, message: str) -> None:
        self.toasts.append(("validation", message))

    def _show_system_error_toast(self, step: int) -> None:
        self.toasts.append(("system", str(step)))

    def _show_step_success_toast(self, step: int) -> None:
        self.toasts.append(("success", str(step)))

    def _can_run_step(self, step_no: int) -> tuple[bool, str]:
        del step_no
        return self.can_run, self.can_run_reason


def _ns(module: _Module, *, open_path: str = "", save_path: str = "") -> dict[str, object]:
    logs: list[tuple] = []
    logger = _Logger()

    def _log_process_message(*args, **kwargs) -> None:
        logs.append((args, kwargs))
        notify = kwargs.get("notify")
        if callable(notify):
            notify("validation error", "error")

    def _start_async_operation(target: _Module, **kwargs) -> None:
        target.started = kwargs

    return {
        "t": lambda key, **kwargs: f"T:{key}",
        "_localized_log_messages": lambda _key: ("USER_SUCCESS", "USER_ERROR"),
        "QFileDialog": type(
            "_QD",
            (),
            {
                "getOpenFileName": staticmethod(lambda *_a, **_k: (open_path, "")),
                "getSaveFileName": staticmethod(lambda *_a, **_k: (save_path, "")),
            },
        ),
        "resolve_dialog_start_path": lambda _app, _default=None: "D:/start.xlsx",
        "APP_NAME": "FOCUS",
        "CancellationToken": _Token,
        "JobCancelledError": JobCancelledError,
        "ValidationError": ValidationError,
        "AppSystemError": AppSystemError,
        "build_i18n_log_message": lambda key, fallback="": f"I18N:{key}:{fallback}",
        "_publish_status": lambda _m, msg: module.published.append(msg),
        "_start_async_operation": _start_async_operation,
        "log_process_message": _log_process_message,
        "show_toast": lambda _m, body, *, title, level: module.toasts.append((level, body)),
        "_validate_uploaded_filled_marks_workbook": lambda _path: None,
        "_build_final_report_default_name": lambda p: f"final:{p}",
        "_atomic_copy_file": lambda src, out: Path(out),
        "_logger": logger,
        "_logs": logs,
        "_logger_obj": logger,
    }


def test_upload_filled_marks_async_early_returns_for_busy_or_cancelled_dialog() -> None:
    module = _Module()
    ns = _ns(module, open_path="")

    module.state.busy = True
    step2_phase.upload_filled_marks_async(module, ns=ns)
    assert module.started == {}

    module.state.busy = False
    step2_phase.upload_filled_marks_async(module, ns=ns)
    assert module.started == {}


def test_upload_filled_marks_async_success_paths() -> None:
    module = _Module()
    ns = _ns(module, open_path="D:/filled.xlsx")
    step2_phase.upload_filled_marks_async(module, ns=ns)
    cast(Callable[[object], None], module.started["on_success"])(True)
    assert module.filled_marks_done is True
    assert module.filled_marks_outdated is False
    assert module.filled_marks_path == "D:/filled.xlsx"
    assert any("step2_uploaded_filled" in msg for msg in module.published)
    assert module.toasts[-1] == ("success", "2")
    assert module.remembered == ["D:/filled.xlsx"]

    module2 = _Module()
    module2.filled_marks_done = True
    ns2 = _ns(module2, open_path="D:/filled-new.xlsx")
    step2_phase.upload_filled_marks_async(module2, ns=ns2)
    cast(Callable[[object], None], module2.started["on_success"])(True)
    assert module2.final_report_outdated is True
    assert any("step2_changed_filled" in msg for msg in module2.published)


def test_upload_filled_marks_async_failure_and_work_paths() -> None:
    module = _Module()
    ns = _ns(module, open_path="D:/filled.xlsx")
    step2_phase.upload_filled_marks_async(module, ns=ns)
    on_failure = cast(Callable[[Exception], None], module.started["on_failure"])

    on_failure(JobCancelledError("stop"))
    assert any("operation_cancelled" in msg for msg in module.published)
    assert cast(_Logger, ns["_logger_obj"]).info_calls

    on_failure(ValidationError("bad"))
    assert ("validation", "validation error") in module.toasts

    on_failure(AppSystemError("sys"))
    assert ("system", "2") in module.toasts

    on_failure(RuntimeError("oops"))
    assert module.toasts.count(("system", "2")) >= 2

    assert cast(Callable[[], bool], module.started["work"])() is True


def test_generate_final_report_async_early_gate_paths() -> None:
    module = _Module()
    ns = _ns(module)

    module.state.busy = True
    step2_phase.generate_final_report_async(module, ns=ns)
    assert module.started == {}

    module.state.busy = False
    module.can_run = False
    module.can_run_reason = "Need upload"
    step2_phase.generate_final_report_async(module, ns=ns)
    assert module.toasts[-1] == ("info", "Need upload")

    module.can_run = True
    module.filled_marks_done = False
    step2_phase.generate_final_report_async(module, ns=ns)
    assert module.toasts[-1] == ("info", "T:instructor.require.step2")

    module.filled_marks_done = True
    module.filled_marks_outdated = False
    ns_cancel = _ns(module, save_path="")
    step2_phase.generate_final_report_async(module, ns=ns_cancel)
    assert module.started == {}


def test_generate_final_report_async_source_missing_and_success() -> None:
    module = _Module()
    module.filled_marks_done = True
    module.filled_marks_outdated = False
    module.filled_marks_path = "D:/missing.xlsx"
    ns = _ns(module, save_path="D:/final.xlsx")
    step2_phase.generate_final_report_async(module, ns=ns)
    assert cast(_Logger, ns["_logger_obj"]).warning_calls
    assert module.toasts[-1] == ("error", "T:instructor.require.step2")
    assert module.started == {}

    source = Path("tests/.tmp-filled.xlsx")
    source.write_text("ok", encoding="utf-8")
    try:
        module2 = _Module()
        module2.filled_marks_done = True
        module2.filled_marks_outdated = False
        module2.filled_marks_path = str(source)
        ns2 = _ns(module2, save_path="D:/final.xlsx")
        step2_phase.generate_final_report_async(module2, ns=ns2)
        cast(Callable[[object], None], module2.started["on_success"])(Path("D:/final.xlsx"))
        assert module2.final_report_done is True
        assert module2.final_report_outdated is False
        assert module2.final_report_path == "D:/final.xlsx"
        assert any("step2_generated" in msg for msg in module2.published)
        assert module2.toasts[-1] == ("success", "2")
        assert module2.remembered == ["D:/final.xlsx"]
    finally:
        source.unlink(missing_ok=True)


def test_generate_final_report_async_failure_and_work_service_fallback() -> None:
    source = Path("tests/.tmp-filled2.xlsx")
    source.write_text("ok", encoding="utf-8")
    try:
        module = _Module()
        module.filled_marks_done = True
        module.filled_marks_outdated = False
        module.filled_marks_path = str(source)
        ns = _ns(module, save_path="D:/final.xlsx")
        step2_phase.generate_final_report_async(module, ns=ns)
        on_failure = cast(Callable[[Exception], None], module.started["on_failure"])

        on_failure(JobCancelledError("stop"))
        assert any("operation_cancelled" in msg for msg in module.published)

        on_failure(ValidationError("bad"))
        assert ("validation", "validation error") in module.toasts

        on_failure(RuntimeError("boom"))
        assert ("system", "2") in module.toasts

        out = cast(Callable[[], Path], module.started["work"])()
        assert out == Path("D:/final.xlsx")

        module2 = _Module()
        module2._workflow_service = _WorkflowService()
        module2.filled_marks_done = True
        module2.filled_marks_outdated = False
        module2.filled_marks_path = str(source)
        ns2 = _ns(module2, save_path="D:/final2.xlsx")
        step2_phase.generate_final_report_async(module2, ns=ns2)
        out2 = cast(Callable[[], Path], module2.started["work"])()
        assert out2 == Path("D:/final2.xlsx")
        assert module2._workflow_service.generated == [(str(source), "D:/final2.xlsx")]
    finally:
        source.unlink(missing_ok=True)

