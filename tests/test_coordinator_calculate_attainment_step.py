from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, TypedDict

from common.exceptions import JobCancelledError
from modules.coordinator.steps import calculate_attainment as step


@dataclass
class _State:
    busy: bool = False


@dataclass
class _Signature:
    section: str


@dataclass
class _JobContext:
    job_id: str
    step_id: str


@dataclass
class _Result:
    output_path: Path
    duplicate_reg_count: int
    duplicate_entries: tuple[tuple[str, str, str], ...]
    inner_join_drop_count: int = 0
    inner_join_drop_details: tuple[str, ...] = ()


class _Logger:
    def __init__(self) -> None:
        self.info_calls: list[tuple[str, tuple[Any, ...], dict[str, Any]]] = []

    def info(self, msg: str, *args: object, **kwargs: object) -> None:
        self.info_calls.append((msg, tuple(args), dict(kwargs)))


class _FileDialog:
    def __init__(self, save_path: str) -> None:
        self.save_path = save_path
        self.calls: list[tuple[object, str, str, str]] = []

    def getSaveFileName(
        self,
        parent: object,
        caption: str = "",
        dir: str = "",
        filter: str = "",
    ) -> tuple[str, str]:
        self.calls.append((parent, caption, dir, filter))
        return self.save_path, ""


class _WorkflowService:
    def __init__(self) -> None:
        self.context_payloads: list[dict[str, object]] = []
        self.calculate_calls: list[tuple[list[Path], Path, object, object, object]] = []

    def create_job_context(self, *, step_id: str, payload: dict[str, object]) -> _JobContext:
        self.context_payloads.append({"step_id": step_id, "payload": payload})
        return _JobContext(job_id="ctx-job", step_id=step_id)

    def calculate_attainment(
        self,
        files: list[Path],
        output: Path,
        *,
        context: object,
        cancel_token: object,
        thresholds: object,
    ) -> str:
        self.calculate_calls.append((files, output, context, cancel_token, thresholds))
        return "workflow-service-result"


class _Module:
    def __init__(self) -> None:
        self.state = _State()
        self._files: list[Path] = [Path("C:/in.xlsx")]
        self._downloaded_outputs: list[Path] = []
        self._logger = _Logger()
        self._published: list[tuple[str, dict[str, object]]] = []
        self._remembered_dirs: list[str] = []
        self._started: dict[str, object] = {}
        self._workflow_service: _WorkflowService | None = None
        self._thresholds: tuple[float, float, float] | None = (40.0, 60.0, 75.0)

    def get_attainment_thresholds(self) -> tuple[float, float, float] | None:
        return self._thresholds

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        self._published.append((text_key, kwargs))

    def _remember_dialog_dir_safe(self, path: str) -> None:
        self._remembered_dirs.append(path)

    def _start_async_operation(self, **kwargs: object) -> None:
        self._started = dict(kwargs)

    def _drain_next_batch(self) -> None:
        return None


class _Ctx(TypedDict):
    dialog: _FileDialog
    logs: list[tuple[tuple[object, ...], dict[str, object]]]
    toasts: list[tuple[str, str, str]]
    called: dict[str, list[tuple[list[Path], Path, dict[str, object]]]]


def _make_ns(module: _Module, *, save_path: str = "C:/out.xlsx") -> tuple[dict[str, object], _Ctx]:
    dialog = _FileDialog(save_path)
    logs: list[tuple[tuple[object, ...], dict[str, object]]] = []
    toasts: list[tuple[str, str, str]] = []
    called: dict[str, list[tuple[list[Path], Path, dict[str, object]]]] = {"generate_calls": []}

    def _t(key: str, **kwargs: object) -> str:
        return f"T:{key}"

    def _generate(files: list[Path], output: Path, **kwargs: object) -> str:
        called["generate_calls"].append((files, output, kwargs))
        return "generator-result"

    ns = {
        "t": _t,
        "QFileDialog": dialog,
        "APP_NAME": "COTAS",
        "resolve_dialog_start_path": lambda _app, default_name: f"C:/start/{default_name}",
        "_extract_final_report_signature": lambda _p: _Signature(section="A"),
        "_build_co_attainment_default_name": lambda _p, section="": f"co-{section}.xlsx",
        "_CoAttainmentWorkbookResult": _Result,
        "_path_key": lambda p: str(p).lower(),
        "log_process_message": lambda *args, **kwargs: logs.append((args, kwargs)),
        "build_i18n_log_message": lambda key, **kwargs: f"MSG:{key}",
        "show_toast": lambda _parent, message, *, title, level: toasts.append((message, title, level)),
        "_generate_co_attainment_workbook": _generate,
    }
    return ns, {"dialog": dialog, "logs": logs, "toasts": toasts, "called": called}


def _run_started_callback(module: _Module, name: str, *args: object) -> None:
    cb = module._started.get(name)
    assert callable(cb)
    cb(*args)


def test_calculate_attainment_returns_early_for_busy_or_empty_files() -> None:
    module = _Module()
    module.state.busy = True
    ns, _ctx = _make_ns(module)
    step.calculate_attainment_async(module, ns=ns)
    assert module._started == {}

    module = _Module()
    module._files = []
    ns, _ctx = _make_ns(module)
    step.calculate_attainment_async(module, ns=ns)
    assert module._started == {}


def test_calculate_attainment_returns_when_dialog_cancelled_or_thresholds_missing() -> None:
    module = _Module()
    ns, _ctx = _make_ns(module, save_path="")
    step.calculate_attainment_async(module, ns=ns)
    assert module._started == {}

    module = _Module()
    module._thresholds = None
    ns, _ctx = _make_ns(module)
    step.calculate_attainment_async(module, ns=ns)
    assert module._started == {}


def test_calculate_attainment_success_with_result_object_and_duplicates() -> None:
    module = _Module()
    ns, ctx = _make_ns(module, save_path="C:/generated.xlsx")

    step.calculate_attainment_async(module, ns=ns)
    _run_started_callback(
        module,
        "on_success",
        _Result(
            output_path=Path("C:/generated.xlsx"),
            duplicate_reg_count=2,
            duplicate_entries=(("R1", "Sheet1", "WB1"), ("R2", "Sheet2", "WB2")),
            inner_join_drop_count=1,
            inner_join_drop_details=("secA CO1 dropped=1",),
        ),
    )

    assert module._downloaded_outputs == [Path("C:/generated.xlsx")]
    assert module._remembered_dirs == [str(Path("C:/generated.xlsx"))]
    assert any(key == "coordinator.status.calculate_completed" for key, _ in module._published)
    assert any(key == "coordinator.regno_dedup.log_body" for key, _ in module._published)
    assert len(ctx["toasts"]) == 3
    assert "t:coordinator.join_drop.body" in ctx["toasts"][-1][0].lower()
    logs = ctx["logs"]
    assert logs
    _args, kwargs = logs[-1]
    assert kwargs["success_message"]
    assert kwargs["user_success_message"] == "MSG:coordinator.status.calculate_completed"


def test_calculate_attainment_success_with_string_path_does_not_duplicate_existing() -> None:
    module = _Module()
    module._downloaded_outputs = [Path("C:/OUT.xlsx")]
    ns, _ctx = _make_ns(module, save_path="C:/out.xlsx")

    step.calculate_attainment_async(module, ns=ns)
    _run_started_callback(module, "on_success", "C:/out.xlsx")

    assert module._downloaded_outputs == [Path("C:/OUT.xlsx")]
    assert module._remembered_dirs == [str(Path("C:/out.xlsx"))]


def test_calculate_attainment_failure_paths_cancel_and_error() -> None:
    module = _Module()
    ns, ctx = _make_ns(module)
    step.calculate_attainment_async(module, ns=ns)

    _run_started_callback(module, "on_failure", JobCancelledError("cancel"))
    assert any(key == "coordinator.status.operation_cancelled" for key, _ in module._published)
    assert module._logger.info_calls

    _run_started_callback(module, "on_failure", RuntimeError("boom"))
    assert ctx["logs"]
    _args, kwargs = ctx["logs"][-1]
    assert isinstance(kwargs["error"], RuntimeError)
    assert kwargs["user_error_message"] == "MSG:coordinator.status.processing_failed"
    assert ctx["toasts"][-1] == (
        "T:coordinator.status.processing_failed",
        "T:coordinator.title",
        "error",
    )


def test_calculate_attainment_uses_workflow_service_when_available_otherwise_fallback_generator() -> None:
    module = _Module()
    module._workflow_service = _WorkflowService()
    ns, ctx = _make_ns(module)
    step.calculate_attainment_async(module, ns=ns)

    work = module._started.get("work")
    assert callable(work)
    result = work()
    assert result == "workflow-service-result"
    assert module._workflow_service.calculate_calls
    generate_calls = ctx["called"]["generate_calls"]
    assert not generate_calls

    module2 = _Module()
    ns2, ctx2 = _make_ns(module2)
    step.calculate_attainment_async(module2, ns=ns2)
    work2 = module2._started.get("work")
    assert callable(work2)
    result2 = work2()
    assert result2 == "generator-result"
    generate_calls2 = ctx2["called"]["generate_calls"]
    assert len(generate_calls2) == 1
