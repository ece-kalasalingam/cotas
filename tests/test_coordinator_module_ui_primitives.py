from __future__ import annotations

from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import QRect, Qt
from PySide6.QtWidgets import QApplication, QListWidgetItem

from modules import coordinator_module as coordinator_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module_no_setup_stub(monkeypatch: pytest.MonkeyPatch) -> coordinator_ui.CoordinatorModule:
    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(coordinator_ui, "_setup_ui_logging_impl", lambda _m, ns: None)
    return coordinator_ui.CoordinatorModule()


def test_excel_drop_list_paint_and_event_edges(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    dl = coordinator_ui._ExcelDropList()

    dl.set_placeholder_text("Drop files")

    class _Mime:
        def __init__(self, has_urls: bool, urls: list[object] | None = None) -> None:
            self._has = has_urls
            self._urls = urls or []

        def hasUrls(self):  # noqa: N802
            return self._has

        def urls(self):
            return list(self._urls)

    class _Url:
        def __init__(self, path: str, *, local: bool) -> None:
            self._path = path
            self._local = local

        def isLocalFile(self):  # noqa: N802
            return self._local

        def toLocalFile(self):  # noqa: N802
            return self._path

    class _Evt:
        def __init__(self, mime, button=Qt.MouseButton.LeftButton):
            self._mime = mime
            self._button = button
            self.accepted = 0
            self.ignored = 0

        def mimeData(self):  # noqa: N802
            return self._mime

        def acceptProposedAction(self):  # noqa: N802
            self.accepted += 1

        def ignore(self):
            self.ignored += 1

        def button(self):
            return self._button

        def accept(self):
            self.accepted += 1

    entered = _Evt(_Mime(False))
    dl.dragEnterEvent(cast(Any, entered))
    assert entered.ignored == 1

    moved = _Evt(_Mime(True))
    dl.dragMoveEvent(cast(Any, moved))
    assert moved.accepted == 1

    leave_calls = {"count": 0}
    monkeypatch.setattr(
        coordinator_ui.QListWidget,
        "dragLeaveEvent",
        lambda self, e: leave_calls.__setitem__("count", leave_calls["count"] + 1),
    )
    dl.dragLeaveEvent(cast(Any, _Evt(_Mime(False))))
    assert leave_calls["count"] == 1

    dropped_events: list[tuple[str, tuple[str, ...]]] = []
    dl.files_dropped.connect(lambda paths: dropped_events.append(("drop", tuple(paths))))
    drop_evt = _Evt(_Mime(True, urls=[_Url("Z:/net.xlsx", local=False)]))
    dl.dropEvent(cast(Any, drop_evt))
    assert drop_evt.ignored == 1
    assert dropped_events == []

    dbl_calls = {"count": 0}
    monkeypatch.setattr(
        coordinator_ui.QListWidget,
        "mouseDoubleClickEvent",
        lambda self, e: dbl_calls.__setitem__("count", dbl_calls["count"] + 1),
    )
    dl.mouseDoubleClickEvent(_Evt(_Mime(False), button=Qt.MouseButton.RightButton))
    assert dbl_calls["count"] == 1


def test_drop_zone_and_elided_label_paint_and_resize(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    frame = coordinator_ui._DropZoneFrame()
    frame.setProperty("dragActive", True)
    frame.resize(80, 40)
    frame.update()

    label = coordinator_ui._ElidedFileNameLabel("very_long_name_for_elide.xlsx")

    monkeypatch.setattr(label, "contentsRect", lambda: QRect(0, 0, 0, 10))
    label._apply_elided_text()

    monkeypatch.setattr(label, "contentsRect", lambda: QRect(0, 0, 80, 10))
    label.resizeEvent(None)


def test_file_item_widget_fallback_icon_and_emit(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)

    class _NullIcon:
        def isNull(self) -> bool:  # noqa: N802
            return True

    class _Style:
        def standardIcon(self, *_args, **_kwargs):
            return _NullIcon()

    monkeypatch.setattr(coordinator_ui._CoordinatorFileItemWidget, "style", lambda self: _Style())

    item_widget = coordinator_ui._CoordinatorFileItemWidget("C:/a.xlsx")
    removed: list[str] = []
    item_widget.removed.connect(lambda path: removed.append(path))

    assert item_widget.remove_btn.text()
    item_widget.remove_btn.click()
    assert removed == ["C:/a.xlsx"]


def test_refresh_ui_disables_remove_buttons_and_new_widget_factory(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module_no_setup_stub(monkeypatch)
    module._files = [coordinator_ui.Path("C:/a.xlsx")]

    item = QListWidgetItem()
    module.drop_list.addItem(item)
    row_widget = module._new_file_item_widget("C:/a.xlsx", parent=module.drop_list)
    item.setSizeHint(row_widget.sizeHint())
    module.drop_list.setItemWidget(item, row_widget)

    module.state.busy = True
    module._refresh_ui()
    assert row_widget.remove_btn.isEnabled() is False

    module.state.busy = False
    module._refresh_ui()
    assert row_widget.remove_btn.isEnabled() is True
    module.close()


def test_setup_ui_logging_wrapper_invokes_impl(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    seen = {"count": 0}

    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(
        coordinator_ui,
        "_setup_ui_logging_impl",
        lambda _m, ns: seen.__setitem__("count", seen["count"] + 1),
    )

    module = coordinator_ui.CoordinatorModule()
    assert seen["count"] >= 1
    module.close()


def test_excel_drop_list_paint_returns_when_no_placeholder(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    dl = coordinator_ui._ExcelDropList()
    dl.set_placeholder_text("")
    assert dl._placeholder_text == ""
