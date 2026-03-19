from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from common.jobs import CancellationToken
from modules import coordinator_module as coordinator_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module(monkeypatch: pytest.MonkeyPatch) -> coordinator_ui.CoordinatorModule:
    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(coordinator_ui.CoordinatorModule, "_setup_ui_logging", lambda self: None)
    return coordinator_ui.CoordinatorModule()


def test_processing_wrapper_uses_underlying_validator(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(coordinator_ui, "_processing_has_valid_final_co_report", lambda p: str(p).endswith("ok.xlsx"))
    assert coordinator_ui._has_valid_final_co_report(Path("ok.xlsx")) is True
    assert coordinator_ui._has_valid_final_co_report(Path("bad.xlsx")) is False


def test_publish_and_delegate_helpers(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    seen: list[tuple[str, tuple, dict]] = []

    monkeypatch.setattr(module._runtime, "publish_status", lambda message: seen.append(("status", (message,), {})))
    monkeypatch.setattr(module._runtime, "publish_status_key", lambda text_key, **kwargs: seen.append(("key", (text_key,), kwargs)))
    monkeypatch.setattr(coordinator_ui, "calculate_attainment_async", lambda _m, ns: seen.append(("calc", (), {})))
    monkeypatch.setattr(coordinator_ui, "process_files_async", lambda _m, dropped, ns: seen.append(("process", (tuple(dropped),), {})))

    module._publish_status("hello")
    module._publish_status_key("coordinator.status.added", count=1)
    module._on_calculate_clicked()
    module._process_files_async(["a.xlsx"])

    assert ("status", ("hello",), {}) in seen
    assert any(tag == "key" and args == ("coordinator.status.added",) for tag, args, _ in seen)
    assert any(tag == "calc" for tag, *_ in seen)
    assert any(tag == "process" for tag, *_ in seen)
    module.close()


def test_async_and_output_link_wrappers_delegate(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    seen: list[str] = []

    class _Runner:
        def start(self, **kwargs):
            seen.append("runner")

    module._async_runner = cast(Any, _Runner())

    monkeypatch.setattr(coordinator_ui, "_add_uploaded_paths_impl", lambda _m, paths, ns: seen.append(f"add:{len(paths)}"))
    monkeypatch.setattr(coordinator_ui, "_output_link_markup_impl", lambda _m, label, path, ns: f"M:{label}:{path}")
    monkeypatch.setattr(coordinator_ui, "_output_links_html_impl", lambda _m, ns: "HTML")
    monkeypatch.setattr(coordinator_ui, "_refresh_output_links_impl", lambda _m, ns: seen.append("refresh"))
    monkeypatch.setattr(coordinator_ui, "_on_output_link_activated_impl", lambda _m, href, ns: seen.append(f"open:{href}"))

    module._start_async_operation(
        token=CancellationToken(),
        job_id="j",
        work=lambda: None,
        on_success=lambda _r: None,
        on_failure=lambda _e: None,
    )
    module._add_uploaded_paths([Path("a.xlsx")])
    assert module._output_link_markup("L", "P") == "M:L:P"
    assert module._output_links_html() == "HTML"
    module._refresh_output_links()
    module._on_output_link_activated("file::x")

    assert "runner" in seen
    assert "add:1" in seen
    assert "refresh" in seen
    assert "open:file::x" in seen
    module.close()


def test_on_files_dropped_remembers_first_path(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    seen: dict[str, object] = {"remember": None, "process": None}
    monkeypatch.setattr(module, "_remember_dialog_dir_safe", lambda path: seen.__setitem__("remember", path))
    monkeypatch.setattr(module, "_process_files_async", lambda dropped: seen.__setitem__("process", list(dropped)))

    module._on_files_dropped(["C:/x.xlsx", "C:/y.xlsx"])

    assert seen["remember"] == "C:/x.xlsx"
    assert seen["process"] == ["C:/x.xlsx", "C:/y.xlsx"]
    module.close()


def test_clear_info_selection_and_drop_active(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    for view in (module.user_log_view, module.generated_outputs_view):
        if hasattr(view, "setPlainText"):
            view.setPlainText("hello")
        else:
            cast(Any, view).setText("hello")
        cursor = view.textCursor()
        cursor.select(cursor.SelectionType.Document)
        view.setTextCursor(cursor)

    module._clear_info_text_selection()
    assert all(not v.textCursor().hasSelection() for v in (module.user_log_view, module.generated_outputs_view))

    updates = {"count": 0}
    monkeypatch.setattr(module.drop_zone, "update", lambda: updates.__setitem__("count", updates["count"] + 1))
    module._set_drop_active(True)
    assert bool(module.drop_zone.property("dragActive")) is True
    assert updates["count"] == 1
    module.close()


def test_drop_list_event_branches(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    dl = coordinator_ui._ExcelDropList()
    events: list[tuple[str, object]] = []
    dl.drag_state_changed.connect(lambda v: events.append(("drag", v)))
    dl.browse_requested.connect(lambda: events.append(("browse", True)))
    dl.files_dropped.connect(lambda paths: events.append(("drop", tuple(paths))))

    class _Mime:
        def __init__(self, has_urls: bool, urls: list[object] | None = None) -> None:
            self._has = has_urls
            self._urls = urls or []

        def hasUrls(self):  # noqa: N802
            return self._has

        def urls(self):
            return list(self._urls)

    class _Url:
        def __init__(self, path: str, local: bool = True) -> None:
            self._p = path
            self._l = local

        def isLocalFile(self):  # noqa: N802
            return self._l

        def toLocalFile(self):  # noqa: N802
            return self._p

    class _Evt:
        def __init__(self, mime):
            self._mime = mime
            self.accepted = 0
            self.ignored = 0
            self._btn = Qt.MouseButton.LeftButton

        def mimeData(self):  # noqa: N802
            return self._mime

        def acceptProposedAction(self):  # noqa: N802
            self.accepted += 1

        def ignore(self):
            self.ignored += 1

        def button(self):
            return self._btn

        def accept(self):
            self.accepted += 1

    e1 = _Evt(_Mime(True))
    dl.dragEnterEvent(cast(Any, e1))
    assert e1.accepted == 1

    e2 = _Evt(_Mime(False))
    dl.dragMoveEvent(cast(Any, e2))
    assert e2.ignored == 1

    e3 = _Evt(_Mime(True, urls=[_Url("C:/a.xlsx"), _Url("Z:/n", local=False)]))
    dl.dropEvent(cast(Any, e3))
    assert any(tag == "drop" and payload == ("C:/a.xlsx",) for tag, payload in events)

    e4 = _Evt(_Mime(False))
    dl.mouseDoubleClickEvent(e4)
    assert any(tag == "browse" for tag, _ in events)


def test_on_files_dropped_skips_remember_when_all_paths_empty(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    seen: dict[str, object] = {"remember": 0, "process": None}

    monkeypatch.setattr(
        module,
        "_remember_dialog_dir_safe",
        lambda _path: seen.__setitem__("remember", cast(int, seen["remember"]) + 1),
    )
    monkeypatch.setattr(module, "_process_files_async", lambda dropped: seen.__setitem__("process", list(dropped)))

    module._on_files_dropped(["", ""])

    assert seen["remember"] == 0
    assert seen["process"] == ["", ""]
    module.close()


def test_misc_wrapper_methods_delegate_and_toggle(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    seen: list[str] = []

    monkeypatch.setattr(module._runtime, "setup_ui_logging", lambda: seen.append("setup"))
    monkeypatch.setattr(module._runtime, "append_user_log", lambda message: seen.append(f"append:{message}"))
    monkeypatch.setattr(coordinator_ui, "_rerender_user_log_impl", lambda _m, ns: seen.append("rerender"))
    monkeypatch.setattr(module, "_output_links_html", lambda: "outputs")
    monkeypatch.setattr(coordinator_ui, "remove_file_by_path", lambda _m, path, ns: seen.append(f"remove:{path}"))
    monkeypatch.setattr(coordinator_ui, "clear_all", lambda _m, ns: seen.append("clear"))

    module._setup_ui_logging()
    module._append_user_log("x")
    module._rerender_user_log()
    module.set_shared_activity_log_mode(True)
    assert module.info_tabs.isHidden() is True
    module.set_shared_activity_log_mode(False)
    assert module.info_tabs.isHidden() is False
    assert module.get_shared_outputs_html() == "outputs"
    module._remove_file_by_path("C:/a.xlsx")
    module._clear_all()

    assert "append:x" in seen
    assert "rerender" in seen
    assert "remove:C:/a.xlsx" in seen
    assert "clear" in seen
    module.close()


def test_set_busy_calls_state_and_refresh(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"refresh": 0}
    monkeypatch.setattr(module, "_refresh_ui", lambda: calls.__setitem__("refresh", calls["refresh"] + 1))

    module._set_busy(True, job_id="job-1")
    assert module.state.busy is True
    assert module.state.active_job_id == "job-1"
    assert calls["refresh"] == 1

    module._set_busy(False)
    assert module.state.busy is False
    assert module.state.active_job_id is None
    assert calls["refresh"] == 2
    module.close()


def test_get_attainment_thresholds_enforces_strict_bounds(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    toasts: list[tuple[str, str, str]] = []
    status_keys: list[str] = []
    monkeypatch.setattr(
        coordinator_ui,
        "show_toast",
        lambda _parent, message, *, title, level: toasts.append((message, title, level)),
    )
    monkeypatch.setattr(module, "_publish_status_key", lambda key, **_kwargs: status_keys.append(key))

    module.threshold_l1_input.setValue(50.0)
    module.threshold_l2_input.setValue(70.0)
    module.threshold_l3_input.setValue(90.0)
    assert module.get_attainment_thresholds() == (50.0, 70.0, 90.0)

    module.threshold_l1_input.setValue(0.0)
    assert module.get_attainment_thresholds() is None
    assert toasts[-1] == (
        coordinator_ui.CoordinatorModule._THRESHOLD_VALIDATION_KEY,
        "coordinator.title",
        "error",
    )
    assert status_keys[-1] == coordinator_ui.CoordinatorModule._THRESHOLD_VALIDATION_KEY
    module.close()

