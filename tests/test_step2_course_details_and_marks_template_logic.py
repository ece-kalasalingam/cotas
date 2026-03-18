from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, cast

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
    @dataclass
    class _State:
        busy: bool = False

    def __init__(self) -> None:
        self.state = _Module._State()
        self.marks_template_done = False
        self.step2_upload_ready = False
        self.step2_course_details_path: str | None = None
        self.marks_template_path: str | None = None
        self.filled_marks_done = False
        self.final_report_done = False
        self.filled_marks_outdated = False
        self.final_report_outdated = False
        self._step2_marks_default_name = "marks_template.xlsx"
        self._workflow_service: Any = None
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


def _ns(module: _Module, *, open_path: str = "", save_path: str = "") -> dict[str, object]:
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
        "_publish_status_compat": lambda _m, msg: module.published.append(msg),
        "log_process_message": _log_process_message,
        "_logger": logger,
        "_start_async_operation_compat": _start_async_operation_compat,
        "show_toast": lambda _m, body, *, title, level: module.toasts.append((level, body)),
        "validate_course_details_workbook": lambda p: None,
        "_build_marks_template_default_name": lambda p: f"DEFAULT:{p}",
        "generate_marks_template_from_course_details": lambda _src, _out: None,
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
    module.marks_template_done = True
    module.filled_marks_done = True
    module.final_report_done = True
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    cast(Callable[[object], None], module.started["on_success"])({"default_marks_name": "X.xlsx"})

    assert module.step2_course_details_path == "D:/course.xlsx"
    assert module.step2_upload_ready is True
    assert module.marks_template_done is False
    assert module.marks_template_path is None
    assert module._step2_marks_default_name == "X.xlsx"
    assert module.filled_marks_outdated is True
    assert module.final_report_outdated is True
    assert any("step1_changed" in msg for msg in module.published)
    assert ("success", "1") not in module.toasts


def test_upload_course_details_async_failure_branches() -> None:
    module = _Module()
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    on_failure = cast(Callable[[Exception], None], module.started["on_failure"])

    on_failure(JobCancelledError("stop"))
    assert any("operation_cancelled" in msg for msg in module.published)
    assert cast(_Logger, ns["_logger_obj"]).info_calls

    on_failure(ValidationError("bad"))
    assert ("validation", "validation error") in module.toasts

    on_failure(AppSystemError("sys"))
    assert ("system", "1") in module.toasts

    on_failure(RuntimeError("oops"))
    assert module.toasts.count(("system", "1")) >= 2


def test_upload_course_details_async_work_uses_service_and_fallback() -> None:
    module = _Module()
    service = _WorkflowService()
    module._workflow_service = service
    ns = _ns(module, open_path="D:/course.xlsx")

    step2.upload_course_details_async(module, ns=ns)
    result = cast(Callable[[], dict[str, str]], module.started["work"])()

    assert service.validated == ["D:/course.xlsx"]
    assert result["default_marks_name"] == "DEFAULT:D:/course.xlsx"

    module2 = _Module()
    called = {"validate": 0}
    ns2 = _ns(module2, open_path="D:/course2.xlsx")
    ns2["validate_course_details_workbook"] = lambda _p: called.__setitem__("validate", called["validate"] + 1)
    step2.upload_course_details_async(module2, ns=ns2)
    cast(Callable[[], dict[str, str]], module2.started["work"])()
    assert called["validate"] == 1


def test_prepare_marks_template_async_early_returns_and_prereq_toast() -> None:
    module = _Module()
    ns = _ns(module, save_path="")

    module.state.busy = True
    step2.prepare_marks_template_async(module, ns=ns)
    assert module.started == {}

    module.state.busy = False
    module.step2_upload_ready = False
    module.step2_course_details_path = None
    step2.prepare_marks_template_async(module, ns=ns)
    assert module.toasts[-1] == ("info", "T:instructor.require.step1")
    assert module.started == {}


def test_prepare_marks_template_async_returns_when_save_dialog_cancelled() -> None:
    module = _Module()
    module.step2_upload_ready = True
    module.step2_course_details_path = "D:/course.xlsx"
    ns = _ns(module, save_path="")
    step2.prepare_marks_template_async(module, ns=ns)
    assert module.started == {}


def test_prepare_marks_template_async_success_sets_paths_and_outdated_flags() -> None:
    module = _Module()
    module.step2_upload_ready = True
    module.step2_course_details_path = "D:/course.xlsx"
    module.filled_marks_done = True
    module.final_report_done = True
    ns = _ns(module, save_path="D:/marks.xlsx")

    step2.prepare_marks_template_async(module, ns=ns)
    cast(Callable[[object], None], module.started["on_success"])(object())

    assert module.marks_template_done is True
    assert module.marks_template_path == "D:/marks.xlsx"
    assert module.filled_marks_outdated is True
    assert module.final_report_outdated is True
    assert any("step1_prepared" in msg for msg in module.published)
    assert module.toasts[-1] == ("success", "1")


def test_prepare_marks_template_async_failure_branches() -> None:
    module = _Module()
    module.step2_upload_ready = True
    module.step2_course_details_path = "D:/course.xlsx"
    ns = _ns(module, save_path="D:/marks.xlsx")

    step2.prepare_marks_template_async(module, ns=ns)
    on_failure = cast(Callable[[Exception], None], module.started["on_failure"])

    on_failure(JobCancelledError("stop"))
    assert any("operation_cancelled" in msg for msg in module.published)
    assert cast(_Logger, ns["_logger_obj"]).info_calls

    on_failure(ValidationError("bad"))
    assert ("validation", "validation error") in module.toasts

    on_failure(AppSystemError("sys"))
    assert ("system", "1") in module.toasts

    on_failure(RuntimeError("oops"))
    assert module.toasts.count(("system", "1")) >= 2


def test_prepare_marks_template_async_work_uses_service_and_fallback() -> None:
    module = _Module()
    module.step2_upload_ready = True
    module.step2_course_details_path = "D:/course.xlsx"
    module._workflow_service = type(
        "_Service",
        (),
        {
            "create_job_context": staticmethod(lambda *, step_id, payload: _Ctx(job_id="job-2", step_id=step_id)),
            "generate_marks_template": staticmethod(
                lambda source, output, *, context, cancel_token: Path(output)
            ),
        },
    )()
    ns = _ns(module, save_path="D:/marks.xlsx")
    step2.prepare_marks_template_async(module, ns=ns)
    out = cast(Callable[[], Path], module.started["work"])()
    assert out == Path("D:/marks.xlsx")

    module2 = _Module()
    module2.step2_upload_ready = True
    module2.step2_course_details_path = "D:/course2.xlsx"
    called = {"gen": 0}
    ns2 = _ns(module2, save_path="D:/marks2.xlsx")
    ns2["generate_marks_template_from_course_details"] = lambda _src, _out: called.__setitem__(
        "gen", called["gen"] + 1
    )
    step2.prepare_marks_template_async(module2, ns=ns2)
    out2 = cast(Callable[[], Path], module2.started["work"])()
    assert out2 == Path("D:/marks2.xlsx")
    assert called["gen"] == 1
