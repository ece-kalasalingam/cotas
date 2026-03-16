from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from common import toast


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_resolve_parent_returns_none_when_no_qapp(monkeypatch: pytest.MonkeyPatch) -> None:
    class _NoApp:
        @staticmethod
        def instance():
            return None

    monkeypatch.setattr(toast, "QApplication", _NoApp)
    assert toast._resolve_parent(None) is None


def test_resolve_parent_returns_none_when_all_windows_hidden(monkeypatch: pytest.MonkeyPatch) -> None:
    class _W:
        def __init__(self, visible: bool, minimized: bool) -> None:
            self._v = visible
            self._m = minimized

        def isVisible(self):  # noqa: N802
            return self._v

        def isMinimized(self):  # noqa: N802
            return self._m

    class _App:
        @staticmethod
        def instance():
            return _App()

        @staticmethod
        def activeWindow():  # noqa: N802
            return _W(False, False)

        @staticmethod
        def topLevelWidgets():  # noqa: N802
            return [_W(False, False), _W(True, True)]

    monkeypatch.setattr(toast, "QApplication", _App)
    assert toast._resolve_parent(None) is None


def test_toast_widget_fit_width_toggles_word_wrap(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    monkeypatch.setattr(toast, "_IS_WINDOWS", True)
    w = toast._ToastWidget(None, title="Title", message="a very very very long line", level="info")
    try:
        w.fit_width(80)
        assert w._body_label is not None
        assert w._body_label.wordWrap() is True

        w.fit_width(1000)
        assert w._body_label.wordWrap() is False
    finally:
        w.close()


def test_toast_widget_non_windows_sets_shadow(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    monkeypatch.setattr(toast, "_IS_WINDOWS", False)
    w = toast._ToastWidget(None, title="", message="hello", level="success")
    try:
        assert w.graphicsEffect() is not None
    finally:
        w.close()


def test_show_toast_uses_adjust_size_when_available_width_non_positive(monkeypatch: pytest.MonkeyPatch) -> None:
    class _Host:
        def isVisible(self):  # noqa: N802
            return True

        def isMinimized(self):  # noqa: N802
            return False

        def width(self):
            return 0

        def mapToGlobal(self, _point):  # noqa: N802
            class _P:
                @staticmethod
                def x():
                    return 0

                @staticmethod
                def y():
                    return 0

            return _P()

    class _StubToast:
        def __init__(self, *_args, **_kwargs):
            self.adjust_count = 0
            self.fit_calls = []
            self.moves = []

        def width(self):
            return 120

        def fit_width(self, value):
            self.fit_calls.append(value)

        def adjustSize(self):  # noqa: N802
            self.adjust_count += 1

        def move(self, x, y):
            self.moves.append((x, y))

        def show(self):
            return None

        def close(self):
            return None

    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: _Host())
    monkeypatch.setattr(toast, "_ToastWidget", _StubToast)
    monkeypatch.setattr(toast.QTimer, "singleShot", lambda *_args, **_kwargs: None)

    toast.show_toast(None, "x")
