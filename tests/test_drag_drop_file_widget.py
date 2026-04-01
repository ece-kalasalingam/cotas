from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from common.drag_drop_file_widget import DragDropFileList


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    """Qapp.
    
    Args:
        None.
    
    Returns:
        QApplication: Return value.
    
    Raises:
        None.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


class _Mime:
    def __init__(self, has_urls: bool, urls: list[object] | None = None) -> None:
        """Init.
        
        Args:
            has_urls: Parameter value (bool).
            urls: Parameter value (list[object] | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._has = has_urls
        self._urls = urls or []

    def hasUrls(self):  # noqa: N802
        """Hasurls.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._has

    def urls(self):
        """Urls.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return list(self._urls)


class _Url:
    def __init__(self, path: str, *, local: bool = True) -> None:
        """Init.
        
        Args:
            path: Parameter value (str).
            local: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._path = path
        self._local = local

    def isLocalFile(self):  # noqa: N802
        """Islocalfile.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._local

    def toLocalFile(self):  # noqa: N802
        """Tolocalfile.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._path


class _Evt:
    def __init__(self, mime, button=Qt.MouseButton.LeftButton):
        """Init.
        
        Args:
            mime: Parameter value.
            button: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._mime = mime
        self._button = button
        self.accepted = 0
        self.ignored = 0

    def mimeData(self):  # noqa: N802
        """Mimedata.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._mime

    def acceptProposedAction(self):  # noqa: N802
        """Acceptproposedaction.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.accepted += 1

    def ignore(self):
        """Ignore.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.ignored += 1

    def button(self):
        """Button.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._button

    def accept(self):
        """Accept.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.accepted += 1


def test_drag_drop_file_list_multiple_mode_emits_all_local_paths(qapp: QApplication) -> None:
    """Test drag drop file list multiple mode emits all local paths.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Test drag drop file list single mode emits first local path only.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
