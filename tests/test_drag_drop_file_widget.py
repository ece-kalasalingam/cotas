from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from common.drag_drop_file_widget import DragDropFileList


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


class _Mime:
    def __init__(self, has_urls: bool, urls: list[object] | None = None) -> None:
        self._has = has_urls
        self._urls = urls or []

    def hasUrls(self):  # noqa: N802
        return self._has

    def urls(self):
        return list(self._urls)


class _Url:
    def __init__(self, path: str, *, local: bool = True) -> None:
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


def test_drag_drop_file_list_multiple_mode_emits_all_local_paths(qapp: QApplication) -> None:
    dl = DragDropFileList(
        placeholder_margins=(0, 0, 0, 0),
        drop_mode="multiple",
    )
    seen: list[list[str]] = []
    dl.files_dropped.connect(lambda paths: seen.append(list(paths)))

    evt = _Evt(
        _Mime(
            True,
            urls=[
                _Url("C:/a.xlsx", local=True),
                _Url("C:/b.xlsx", local=True),
                _Url("X:/net.xlsx", local=False),
            ],
        )
    )
    dl.dropEvent(cast(Any, evt))
    assert seen == [["C:/a.xlsx", "C:/b.xlsx"]]


def test_drag_drop_file_list_single_mode_emits_first_local_path_only(qapp: QApplication) -> None:
    dl = DragDropFileList(
        placeholder_margins=(0, 0, 0, 0),
        drop_mode="single",
    )
    seen: list[list[str]] = []
    dl.files_dropped.connect(lambda paths: seen.append(list(paths)))

    evt = _Evt(
        _Mime(
            True,
            urls=[
                _Url("C:/first.xlsx", local=True),
                _Url("C:/second.xlsx", local=True),
            ],
        )
    )
    dl.dropEvent(cast(Any, evt))
    assert seen == [["C:/first.xlsx"]]
