from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from modules.coordinator import file_actions


class _Item:
    def __init__(self, path: str) -> None:
        self._path = path

    def data(self, _role):
        return self._path


class _DropList:
    def __init__(self, paths: list[str]) -> None:
        self._items = [_Item(p) for p in paths]
        self.cleared = False

    def count(self) -> int:
        return len(self._items)

    def item(self, row: int):
        return self._items[row]

    def takeItem(self, row: int):  # noqa: N802
        return self._items.pop(row)

    def clear(self) -> None:
        self.cleared = True
        self._items.clear()


class _Module:
    def __init__(self, files: list[str]) -> None:
        self.state = SimpleNamespace(busy=False)
        self._files = [Path(p) for p in files]
        self.drop_list = _DropList(files)
        self._logger = object()
        self.published: list[tuple[str, dict]] = []
        self.refreshes = 0

    def _publish_status_key(self, key: str, **kwargs) -> None:
        self.published.append((key, kwargs))

    def _refresh_ui(self) -> None:
        self.refreshes += 1


class _Qt:
    class ItemDataRole:
        UserRole = object()


def _ns(log_calls: list[tuple]):
    return {
        "_path_key": lambda p: str(p).lower(),
        "Qt": _Qt,
        "log_process_message": lambda *args, **kwargs: log_calls.append((args, kwargs)),
        "build_i18n_log_message": lambda *args, **kwargs: "payload",
        "t": lambda key, **kwargs: f"T:{key}",
    }


def test_remove_file_by_path_noops_when_busy_or_not_found() -> None:
    logs: list[tuple] = []
    module = _Module(["C:/a.xlsx"])
    ns = _ns(logs)

    module.state.busy = True
    file_actions.remove_file_by_path(module, "C:/a.xlsx", ns=ns)
    assert len(module._files) == 1

    module.state.busy = False
    file_actions.remove_file_by_path(module, "C:/missing.xlsx", ns=ns)
    assert len(module._files) == 1
    assert module.refreshes == 0
    assert module.published == []
    assert logs == []


def test_remove_file_by_path_removes_matching_file_and_logs() -> None:
    logs: list[tuple] = []
    module = _Module(["C:/a.xlsx", "C:/b.xlsx"])
    ns = _ns(logs)

    file_actions.remove_file_by_path(module, "C:/a.xlsx", ns=ns)

    assert len(module._files) == 1
    assert module._files[0].name == "b.xlsx"
    assert module.drop_list.count() == 1
    assert module.refreshes == 1
    assert module.published == [("coordinator.status.removed", {"count": 1})]
    assert len(logs) == 1


def test_clear_all_noops_when_busy_or_empty_and_clears_when_non_empty() -> None:
    logs: list[tuple] = []
    module = _Module(["C:/a.xlsx", "C:/b.xlsx"])
    ns = _ns(logs)

    module.state.busy = True
    file_actions.clear_all(module, ns=ns)
    assert len(module._files) == 2

    module.state.busy = False
    file_actions.clear_all(module, ns=ns)
    assert module._files == []
    assert module.drop_list.cleared is True
    assert module.refreshes == 1
    assert module.published[-1] == ("coordinator.status.cleared", {"count": 2})
    assert len(logs) == 1

    # empty no-op
    file_actions.clear_all(module, ns=ns)
    assert len(logs) == 1
