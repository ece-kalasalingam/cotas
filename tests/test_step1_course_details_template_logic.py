from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, cast

from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from modules.instructor.steps import step1_course_details_template as step1


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
        self.step1_path: str | None = None
        self.step1_done = False
        self._workflow_service: _WorkflowService | None = None
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
        self.generated: list[str] = []

    def create_job_context(self, *, step_id: str, payload: dict[str, object]) -> _Ctx:
        del payload
        return _Ctx(job_id="job-s1", step_id=step_id)

    def generate_course_details_template(self, path: str, *, context: _Ctx, cancel_token: _Token) -> Path:
        del context, cancel_token
        self.generated.append(path)
        return Path(path)


def _ns(module: _Module, *, save_path: str = "") -> dict[str, object]:
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
        "ID_COURSE_SETUP": "course-setup-v1",
        "t": lambda key, **kwargs: f"T:{key}",
        "_localized_log_messages": lambda _key: ("USER_SUCCESS", "USER_ERROR"),
        "QFileDialog": type(
            "_QD",
            (),
            {
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
        "generate_course_details_template": lambda _path, *, template_id: None,
        "_logs": logs,
        "_logger_obj": logger,
    }


def test_download_course_template_async_early_returns_for_busy_or_cancelled_dialog() -> None:
    module = _Module()
    ns = _ns(module, save_path="")

    module.state.busy = True
    step1.download_course_template_async(module, ns=ns)
    assert module.started == {}

    module.state.busy = False
    step1.download_course_template_async(module, ns=ns)
    assert module.started == {}


def test_download_course_template_async_success_path() -> None:
    module = _Module()
    ns = _ns(module, save_path="D:/course_template.xlsx")

    step1.download_course_template_async(module, ns=ns)
    cast(Callable[[object], None], module.started["on_success"])(object())

    assert module.step1_path == "D:/course_template.xlsx"
    assert module.step1_done is True
    assert any("step1_selected" in msg for msg in module.published)
    assert module.toasts[-1] == ("success", "1")
    assert module.remembered == ["D:/course_template.xlsx"]


def test_download_course_template_async_failure_branches() -> None:
    module = _Module()
    ns = _ns(module, save_path="D:/course_template.xlsx")

    step1.download_course_template_async(module, ns=ns)
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


def test_download_course_template_async_work_uses_service_and_fallback() -> None:
    module = _Module()
    service = _WorkflowService()
    module._workflow_service = service
    ns = _ns(module, save_path="D:/course_template.xlsx")

    step1.download_course_template_async(module, ns=ns)
    out = cast(Callable[[], Path], module.started["work"])()
    assert out == Path("D:/course_template.xlsx")
    assert service.generated == ["D:/course_template.xlsx"]

    module2 = _Module()
    called = {"gen": 0}
    ns2 = _ns(module2, save_path="D:/course_template2.xlsx")
    ns2["generate_course_details_template"] = lambda _path, *, template_id: called.__setitem__(
        "gen", called["gen"] + 1
    )
    step1.download_course_template_async(module2, ns=ns2)
    out2 = cast(Callable[[], Path], module2.started["work"])()
    assert out2 == Path("D:/course_template2.xlsx")
    assert called["gen"] == 1
