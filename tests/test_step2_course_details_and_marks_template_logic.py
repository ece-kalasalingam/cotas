from __future__ import annotations

from dataclasses import dataclass

import pytest

from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from modules.instructor.steps import step2_course_details_and_marks_template as step2


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

    def info(self, *args, **kwargs) -> None:
        self.info_calls.append((args, kwargs))


class _Module:
    def __init__(self) -> None:
        self.state = type("State", (), {"busy": False})()
        self.step2_done = False
        self.step2_upload_ready = False
        self.step2_course_details_path = None
        self.step2_path = None
        self.step3_done = False
        self.step4_done = False
        self.step3_outdated = False
        self.step4_outdated = False
        self._step2_marks_default_name = "marks_template.xlsx"
        self._workflow_service = None
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


class _WorkflowService:
    def __init__(self) -> None:
        self.validated: list[str] = []

    def create_job_context(self, *, step_id: str, payload: dict[str, object]) -> _Ctx:
        del payload
        return _Ctx(job_id="job-1", step_id=step_id)

    def validate_course_details_workbook(self, path: str, *, context: _Ctx, cancel_token: _Token) -> None:
        del context, cancel_token
        self.validated.append(path)


def _ns(module: _Module, *, open_path: str = "") -> dict[str, object]:
    logs: list[tuple] = []
    logger = _Logger()

    def _log_process_message(*args, **kwargs) -> None:
        logs.append((args, kwargs))
        notify = kwargs.get("notify")
        if callable(notify):
            notify("validation error", "error")

    def _start_async_operation_compat(target: _Module, **kwargs) -> None:
        target.started = kwargs

    return {
        "t": lambda key, **kwargs: f"T:{key}",
        "_localized_log_messages": lambda _key: ("USER_SUCCESS", "USER_ERROR"),
        "QFileDialog": type(
            "_QD",
            (),
            {
                "getOpenFileName": staticmethod(lambda *_a, **_k: (open_path, "")),
            },
        ),
        "resolve_dialog_start_path": lambda _app, _default=None: "D:/start.xlsx",
        "APP_NAME": "FOCUS",
        "CancellationToken": _Token,
        "JobCancelledError": JobCancelledError,
        "ValidationError": ValidationError,
        "AppSystemError": AppSystemError,
        "build_i18n_log_message": lambda key, fallback="": f"I18N:{key}:{fallback}",
        "_publish_status_compat": lambda _m, msg: module.published.append(msg),
        "log_process_message": _log_process_message,
        "_logger": logger,
        "_start_async_operation_compat": _start_async_operation_compat,
        "show_toast": lambda _m, body, *, title, level: module.toasts.append((level, body)),
        "validate_course_details_workbook": lambda p: None,
        "_build_marks_template_default_name": lambda p: f"DEFAULT:{p}",
        "_logs": logs,
        "_logger_obj": logger,
    }


def test_upload_course_details_async_early_returns_for_busy_or_cancelled_dialog() -> None:
    module = _Module()
    ns = _ns(module, open_path="")

    module.state.busy = True
    step2.upload_course_details_async(module, ns=ns)
    assert module.started == {}

    module.state.busy = False
    step2.upload_course_details_async(module, ns=ns)
    assert module.started == {}


def test_upload_course_details_async_success_replacing_marks_downstream_outdated() -> None:
    module = _Module()
    module.step2_done = True
    module.step3_done = True
    module.step4_done = True
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    module.started["on_success"]({"default_marks_name": "X.xlsx"})

    assert module.step2_course_details_path == "D:/course.xlsx"
    assert module.step2_upload_ready is True
    assert module.step2_done is False
    assert module.step2_path is None
    assert module._step2_marks_default_name == "X.xlsx"
    assert module.step3_outdated is True
    assert module.step4_outdated is True
    assert any("step2_changed" in msg for msg in module.published)
    assert module.toasts[-1] == ("success", "2")


def test_upload_course_details_async_failure_branches() -> None:
    module = _Module()
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    on_failure = module.started["on_failure"]

    on_failure(JobCancelledError("stop"))
    assert any("operation_cancelled" in msg for msg in module.published)
    assert ns["_logger_obj"].info_calls

    on_failure(ValidationError("bad"))
    assert ("validation", "validation error") in module.toasts

    on_failure(AppSystemError("sys"))
    assert ("system", "2") in module.toasts

    on_failure(RuntimeError("oops"))
    assert module.toasts.count(("system", "2")) >= 2


def test_upload_course_details_async_work_uses_service_and_fallback() -> None:
    module = _Module()
    service = _WorkflowService()
    module._workflow_service = service
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    result = module.started["work"]()

    assert service.validated == ["D:/course.xlsx"]
    assert result["default_marks_name"] == "DEFAULT:D:/course.xlsx"

    module2 = _Module()
    called = {"validate": 0}
    ns2 = _ns(module2, open_path="D:/course2.xlsx")
    ns2["validate_course_details_workbook"] = lambda _p: called.__setitem__("validate", called["validate"] + 1)
    step2.upload_course_details_async(module2, ns=ns2)
    module2.started["work"]()
    assert called["validate"] == 1
