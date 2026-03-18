from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, TypedDict, cast

import pytest

from common.exceptions import JobCancelledError
from modules.coordinator.steps import collect_files


class _Logger:
    def __init__(self) -> None:
        self.info_calls: list[tuple[str, tuple, dict]] = []

    def info(self, msg: str, *args, **kwargs) -> None:
        self.info_calls.append((msg, args, kwargs))


@dataclass
class _State:
    busy: bool = False


class _NS(TypedDict):
    t: Callable[..., str]
    _path_key: Callable[[Path], str]
    _analyze_dropped_files: Callable[..., dict[str, object]]
    QListWidgetItem: type[object]
    build_i18n_log_message: Callable[..., str]
    log_process_message: Callable[..., None]
    show_toast: Callable[..., None]
    _log_calls: list[tuple[tuple[object, ...], dict[str, object]]]


class _Module:
    drop_list: Any
    _remove_file_by_path: Callable[..., None]
    _new_file_item_widget: Callable[..., Any]

    def __init__(self) -> None:
        self.state = _State()
        self._files: list[Path] = []
        self._pending_drop_batches: list[list[str]] = []
        self._logger = _Logger()
        self._published: list[tuple[str, dict]] = []
        self._toasts: list[tuple[str, str, str]] = []
        self._started: dict[str, object] = {}

    def _publish_status_key(self, key: str, **kwargs) -> None:
        self._published.append((key, kwargs))

    def _start_async_operation(self, **kwargs) -> None:
        self._started = kwargs

    def _drain_next_batch(self) -> None:
        return None

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        self._files.extend(added_paths)


@dataclass
class _Result:
    output_path: Path
    duplicate_reg_count: int
    duplicate_entries: tuple[tuple[str, str, str], ...]


def _collect_ns(module: _Module) -> _NS:
    toasts = module._toasts
    logs: list[tuple[tuple[object, ...], dict[str, object]]] = []

    def _show_toast(_parent, message: str, *, title: str, level: str) -> None:
        toasts.append((message, title, level))

    def _log_process_message(*args, **kwargs) -> None:
        logs.append((args, kwargs))

    return {
        "t": lambda key, **kwargs: f"T:{key}",
        "_path_key": lambda p: str(p).lower(),
        "_analyze_dropped_files": lambda *_args, **_kwargs: {},
        "QListWidgetItem": object,
        "build_i18n_log_message": lambda *_args, **_kwargs: "payload",
        "log_process_message": _log_process_message,
        "show_toast": _show_toast,
        "_log_calls": logs,
    }


def test_process_files_async_queue_when_busy() -> None:
    module = _Module()
    module.state.busy = True
    ns = _collect_ns(module)

    collect_files.process_files_async(module, ["a.xlsx", "b.xlsx"], ns=ns)

    assert module._pending_drop_batches == [["a.xlsx", "b.xlsx"]]
    assert any(key == "coordinator.status.queued" for key, _ in module._published)


def test_process_files_async_success_and_cancel_paths() -> None:
    module = _Module()
    module._files = [Path("C:/existing.xlsx")]
    ns = _collect_ns(module)

    collect_files.process_files_async(module, ["C:/new.xlsx"], ns=ns)
    assert "on_success" in module._started and "on_failure" in module._started

    on_success = cast(Callable[[object], None], module._started["on_success"])
    on_failure = cast(Callable[[Exception], None], module._started["on_failure"])

    on_success(
        {
            "added": ["C:/added.xlsx"],
            "duplicates": 1,
            "invalid_final_report": ["C:/bad.xlsx"],
            "ignored": 2,
        }
    )

    assert Path("C:/added.xlsx") in module._files
    assert any(key == "coordinator.status.added" for key, _ in module._published)
    assert any(key == "coordinator.status.ignored" for key, _ in module._published)
    assert len(module._toasts) == 2

    on_failure(JobCancelledError("cancelled"))
    assert any(key == "coordinator.status.operation_cancelled" for key, _ in module._published)


def test_process_files_async_on_finished_rejects_unexpected_result_type() -> None:
    module = _Module()
    ns = _collect_ns(module)

    collect_files.process_files_async(module, ["C:/new.xlsx"], ns=ns)
    on_success = cast(Callable[[object], None], module._started["on_success"])

    with pytest.raises(RuntimeError, match="unexpected result type"):
        on_success("not-a-dict")


def test_process_files_async_non_cancel_failure_logs_and_toasts() -> None:
    module = _Module()
    ns = _collect_ns(module)

    collect_files.process_files_async(module, ["C:/new.xlsx"], ns=ns)
    on_failure = cast(Callable[[Exception], None], module._started["on_failure"])

    exc = RuntimeError("boom")
    on_failure(exc)

    assert ns["_log_calls"]
    _args, kwargs = ns["_log_calls"][-1]
    assert kwargs["error"] is exc
    assert kwargs["user_error_message"] == "payload"
    assert module._toasts[-1] == (
        "T:coordinator.status.processing_failed",
        "T:coordinator.title",
        "error",
    )


def test_add_uploaded_paths_updates_list_and_widget_bindings() -> None:
    class _Item:
        def __init__(self) -> None:
            self.tooltip = ""
            self.data_calls: list[tuple[object, object]] = []
            self.size_hint = None

        def setToolTip(self, value: str) -> None:  # noqa: N802
            self.tooltip = value

        def setData(self, role: object, value: object) -> None:  # noqa: N802
            self.data_calls.append((role, value))

        def setSizeHint(self, value: object) -> None:  # noqa: N802
            self.size_hint = value

    class _Signal:
        def __init__(self) -> None:
            self.connected: list[object] = []

        def connect(self, callback: object) -> None:
            self.connected.append(callback)

    class _RowWidget:
        def __init__(self, path: str) -> None:
            self.path = path
            self.removed = _Signal()

        def sizeHint(self) -> tuple[int, int]:  # noqa: N802
            return (10, 5)

    class _DropList:
        def __init__(self) -> None:
            self.added_items: list[_Item] = []
            self.bound_widgets: list[tuple[_Item, _RowWidget]] = []

        def addItem(self, item: _Item) -> None:  # noqa: N802
            self.added_items.append(item)

        def setItemWidget(self, item: _Item, widget: _RowWidget) -> None:  # noqa: N802
            self.bound_widgets.append((item, widget))

    module = _Module()
    module.drop_list = _DropList()
    module._remove_file_by_path = lambda *_args, **_kwargs: None
    module._new_file_item_widget = lambda path, parent=None: _RowWidget(path)
    ns = {"QListWidgetItem": _Item}

    collect_files.add_uploaded_paths(
        module,
        [Path("C:/A.xlsx"), Path("C:/B.xlsx")],
        ns=ns,
    )

    assert module._files == [Path("C:/A.xlsx"), Path("C:/B.xlsx")]
    assert len(module.drop_list.added_items) == 2
    assert len(module.drop_list.bound_widgets) == 2
    assert module.drop_list.added_items[0].tooltip == "C:\\A.xlsx"
    assert module.drop_list.added_items[1].tooltip == "C:\\B.xlsx"
    assert module.drop_list.added_items[0].data_calls
    assert module.drop_list.added_items[1].data_calls
    assert module.drop_list.bound_widgets[0][1].removed.connected == [module._remove_file_by_path]
    assert module.drop_list.bound_widgets[1][1].removed.connected == [module._remove_file_by_path]


def test_process_files_async_ignores_empty_dropped_files() -> None:
    module = _Module()
    ns = _collect_ns(module)

    collect_files.process_files_async(module, [], ns=ns)

    assert module._started == {}
    assert module._pending_drop_batches == []
    assert module._published == []
