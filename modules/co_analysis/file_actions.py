"""CO Analysis file-list action helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Mapping, Protocol, TypedDict, cast


class _State(Protocol):
    busy: bool


class _DropListItem(Protocol):
    def data(self, role: object) -> object:  # noqa: N802 - Qt-style name
        ...


class _DropList(Protocol):
    def count(self) -> int:
        ...

    def item(self, row: int) -> _DropListItem:
        ...

    def takeItem(self, row: int) -> object:  # noqa: N802 - Qt-style name
        ...

    def clear(self) -> None:
        ...


class _Module(Protocol):
    state: _State
    _files: list[Path]
    drop_list: _DropList
    _logger: object

    def _refresh_ui(self) -> None:
        ...

    def _publish_status_key(self, key: str, **kwargs: object) -> None:
        ...


class FileActionsNamespace(TypedDict):
    canonical_path_key: Callable[[Path], str]
    user_role: object
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    t: Callable[..., str]


def remove_file_by_path(module: object, file_path: str, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_Module, module)
    typed_ns = cast(FileActionsNamespace, ns)
    if typed_module.state.busy:
        return
    target_key = typed_ns["canonical_path_key"](Path(file_path))
    before_count = len(typed_module._files)
    typed_module._files = [
        path for path in typed_module._files if typed_ns["canonical_path_key"](path) != target_key
    ]
    if len(typed_module._files) == before_count:
        return
    for row in range(typed_module.drop_list.count()):
        item = typed_module.drop_list.item(row)
        path_value = str(item.data(typed_ns["user_role"]) or "")
        if typed_ns["canonical_path_key"](Path(path_value)) == target_key:
            typed_module.drop_list.takeItem(row)
            break
    typed_module._refresh_ui()
    typed_module._publish_status_key("coordinator.status.removed", count=1)
    typed_ns["log_process_message"](
        "removing selected co analysis files",
        logger=typed_module._logger,
        success_message="removing selected co analysis files completed successfully. removed=1",
        user_success_message=typed_ns["build_i18n_log_message"](
            "coordinator.status.removed",
            kwargs={"count": 1},
            fallback=typed_ns["t"]("coordinator.status.removed", count=1),
        ),
    )


def clear_all(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_Module, module)
    typed_ns = cast(FileActionsNamespace, ns)
    if typed_module.state.busy:
        return
    if not typed_module._files:
        return
    total = len(typed_module._files)
    typed_module._files.clear()
    typed_module.drop_list.clear()
    typed_module._refresh_ui()
    typed_module._publish_status_key("coordinator.status.cleared", count=total)
    typed_ns["log_process_message"](
        "clearing co analysis files",
        logger=typed_module._logger,
        success_message=f"clearing co analysis files completed successfully. removed={total}",
        user_success_message=typed_ns["build_i18n_log_message"](
            "coordinator.status.cleared",
            kwargs={"count": total},
            fallback=typed_ns["t"]("coordinator.status.cleared", count=total),
        ),
    )
