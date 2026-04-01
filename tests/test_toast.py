from __future__ import annotations

from typing import Any, cast

from common import toast
from common.constants import (
    TOAST_DEFAULT_DURATION_MS,
    TOAST_ERROR_DURATION_MS,
    TOAST_MARGIN,
)


class _FakeWidget:
    def __init__(self, *, visible: bool = True, minimized: bool = False, width: int = 300, origin=(50, 40)) -> None:
        """Init.
        
        Args:
            visible: Parameter value (bool).
            minimized: Parameter value (bool).
            width: Parameter value (int).
            origin: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._visible = visible
        self._minimized = minimized
        self._width = width
        self._origin = origin

    def window(self):
        """Window.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self

    def isVisible(self) -> bool:  # noqa: N802 - Qt-style API
        """Isvisible.
        
        Args:
            None.
        
        Returns:
            bool: Return value.
        
        Raises:
            None.
        """
        return self._visible

    def isMinimized(self) -> bool:  # noqa: N802 - Qt-style API
        """Isminimized.
        
        Args:
            None.
        
        Returns:
            bool: Return value.
        
        Raises:
            None.
        """
        return self._minimized

    def width(self) -> int:
        """Width.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._width

    def mapToGlobal(self, _point):  # noqa: N802 - Qt-style API
        """Maptoglobal.
        
        Args:
            _point: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        class _P:
            def __init__(self, x: int, y: int) -> None:
                """Init.
                
                Args:
                    x: Parameter value (int).
                    y: Parameter value (int).
                
                Returns:
                    None.
                
                Raises:
                    None.
                """
                self._x = x
                self._y = y

            def x(self) -> int:
                """X.
                
                Args:
                    None.
                
                Returns:
                    int: Return value.
                
                Raises:
                    None.
                """
                return self._x

            def y(self) -> int:
                """Y.
                
                Args:
                    None.
                
                Returns:
                    int: Return value.
                
                Raises:
                    None.
                """
                return self._y

        return _P(*self._origin)


class _FakeRect:
    def __init__(self, x: int, y: int, w: int, h: int) -> None:
        """Init.
        
        Args:
            x: Parameter value (int).
            y: Parameter value (int).
            w: Parameter value (int).
            h: Parameter value (int).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._x, self._y, self._w, self._h = x, y, w, h

    def x(self) -> int:
        """X.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._x

    def y(self) -> int:
        """Y.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._y

    def width(self) -> int:
        """Width.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._w

    def height(self) -> int:
        """Height.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._h


class _FakeScreen:
    def __init__(self, rect: _FakeRect) -> None:
        """Init.
        
        Args:
            rect: Parameter value (_FakeRect).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._rect = rect

    def availableGeometry(self):  # noqa: N802
        """Availablegeometry.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return self._rect


class _FakeApp:
    _instance: Any = None
    _active: Any = None
    _top: list[Any] = []
    _screen: Any = None

    @staticmethod
    def instance():
        """Instance.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return _FakeApp._instance

    @staticmethod
    def activeWindow():  # noqa: N802
        """Activewindow.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return _FakeApp._active

    @staticmethod
    def topLevelWidgets():  # noqa: N802
        """Toplevelwidgets.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return list(_FakeApp._top)

    @staticmethod
    def primaryScreen():  # noqa: N802
        """Primaryscreen.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return _FakeApp._screen


class _FakeToastWidget:
    instances = []

    def __init__(self, host, *, title: str, message: str, level: str) -> None:
        """Init.
        
        Args:
            host: Parameter value.
            title: Parameter value (str).
            message: Parameter value (str).
            level: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.host = host
        self.title = title
        self.message = message
        self.level = level
        self._width = 180
        self.fit_width_calls = []
        self.adjust_size_calls = 0
        self.moves = []
        self.shown = False
        self.closed = False
        _FakeToastWidget.instances.append(self)

    def width(self) -> int:
        """Width.
        
        Args:
            None.
        
        Returns:
            int: Return value.
        
        Raises:
            None.
        """
        return self._width

    def fit_width(self, value: int) -> None:
        """Fit width.
        
        Args:
            value: Parameter value (int).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.fit_width_calls.append(value)

    def adjustSize(self) -> None:  # noqa: N802
        """Adjustsize.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.adjust_size_calls += 1

    def move(self, x: int, y: int) -> None:
        """Move.
        
        Args:
            x: Parameter value (int).
            y: Parameter value (int).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.moves.append((x, y))

    def show(self) -> None:
        """Show.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.shown = True

    def close(self) -> None:
        """Close.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.closed = True


def test_resolve_parent_prefers_explicit_parent_window() -> None:
    """Test resolve parent prefers explicit parent window.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    parent = _FakeWidget()
    assert toast._resolve_parent(cast(Any, parent)) is parent


def test_resolve_parent_uses_active_visible_window(monkeypatch) -> None:
    """Test resolve parent uses active visible window.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(toast, "QApplication", _FakeApp)
    _FakeApp._instance = _FakeApp()
    active = _FakeWidget(visible=True, minimized=False)
    _FakeApp._active = active
    _FakeApp._top = []

    assert toast._resolve_parent(None) is active


def test_resolve_parent_falls_back_to_first_visible_toplevel(monkeypatch) -> None:
    """Test resolve parent falls back to first visible toplevel.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(toast, "QApplication", _FakeApp)
    _FakeApp._instance = _FakeApp()
    _FakeApp._active = _FakeWidget(visible=False, minimized=False)
    visible_top = _FakeWidget(visible=True, minimized=False)
    _FakeApp._top = [_FakeWidget(visible=False), visible_top]

    assert toast._resolve_parent(None) is visible_top


def test_show_toast_returns_early_for_hidden_host(monkeypatch) -> None:
    """Test show toast returns early for hidden host.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    hidden = _FakeWidget(visible=False)
    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: hidden)
    monkeypatch.setattr(toast, "_ToastWidget", _FakeToastWidget)

    calls = []
    monkeypatch.setattr(toast.QTimer, "singleShot", lambda ttl, cb: calls.append((ttl, cb)))

    _FakeToastWidget.instances.clear()
    toast.show_toast(None, "hello")

    assert _FakeToastWidget.instances == []
    assert calls == []


def test_show_toast_uses_error_default_ttl_and_host_positioning(monkeypatch) -> None:
    """Test show toast uses error default ttl and host positioning.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    host = _FakeWidget(visible=True, minimized=False, width=320, origin=(100, 80))
    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: host)
    monkeypatch.setattr(toast, "_ToastWidget", _FakeToastWidget)

    _FakeApp._screen = _FakeScreen(_FakeRect(0, 0, 1200, 800))
    monkeypatch.setattr(toast, "QApplication", _FakeApp)

    timer_calls = []
    monkeypatch.setattr(toast.QTimer, "singleShot", lambda ttl, cb: timer_calls.append((ttl, cb)))

    _FakeToastWidget.instances.clear()
    toast.show_toast(None, "err", level="error")

    created = _FakeToastWidget.instances[-1]
    assert created.fit_width_calls, "fit_width should be used when host width is available"
    expected_x = host.mapToGlobal(None).x() + host.width() - created.width() - TOAST_MARGIN
    expected_y = host.mapToGlobal(None).y() + TOAST_MARGIN
    assert created.moves[-1] == (max(expected_x, TOAST_MARGIN), max(expected_y, TOAST_MARGIN))
    assert created.shown is True
    assert timer_calls and timer_calls[-1][0] == TOAST_ERROR_DURATION_MS


def test_show_toast_prefers_custom_duration_over_defaults(monkeypatch) -> None:
    """Test show toast prefers custom duration over defaults.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    host = _FakeWidget(visible=True, minimized=False)
    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: host)
    monkeypatch.setattr(toast, "_ToastWidget", _FakeToastWidget)
    _FakeApp._screen = _FakeScreen(_FakeRect(0, 0, 1200, 800))
    monkeypatch.setattr(toast, "QApplication", _FakeApp)

    timer_calls = []
    monkeypatch.setattr(toast.QTimer, "singleShot", lambda ttl, cb: timer_calls.append((ttl, cb)))

    _FakeToastWidget.instances.clear()
    toast.show_toast(None, "info", level="info", duration_ms=1234)

    assert timer_calls[-1][0] == 1234


def test_show_toast_uses_global_screen_position_when_no_host(monkeypatch) -> None:
    """Test show toast uses global screen position when no host.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: None)
    monkeypatch.setattr(toast, "_ToastWidget", _FakeToastWidget)
    _FakeApp._screen = _FakeScreen(_FakeRect(10, 20, 900, 700))
    monkeypatch.setattr(toast, "QApplication", _FakeApp)

    timer_calls = []
    monkeypatch.setattr(toast.QTimer, "singleShot", lambda ttl, cb: timer_calls.append((ttl, cb)))

    _FakeToastWidget.instances.clear()
    toast.show_toast(None, "hello")

    created = _FakeToastWidget.instances[-1]
    expected_x = 10 + 900 - created.width() - TOAST_MARGIN
    expected_y = 20 + TOAST_MARGIN
    assert created.moves[-1] == (max(expected_x, TOAST_MARGIN), max(expected_y, TOAST_MARGIN))
    assert timer_calls[-1][0] == TOAST_DEFAULT_DURATION_MS


def test_toast_fit_width_early_return_and_no_screen_fallback(monkeypatch) -> None:
    # fit_width early return branch when body label is missing
    """Test toast fit width early return and no screen fallback.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    toast_widget = toast._ToastWidget.__new__(toast._ToastWidget)  # bypass __init__
    toast_widget._body_label = None
    toast_widget._message = ""
    toast._ToastWidget.fit_width(toast_widget, 100)

    monkeypatch.setattr(toast, "_resolve_parent", lambda _p: None)
    monkeypatch.setattr(toast, "_ToastWidget", _FakeToastWidget)
    monkeypatch.setattr(
        toast,
        "QApplication",
        type("_NoScreenApp", (), {"primaryScreen": staticmethod(lambda: None), "instance": staticmethod(lambda: None)}),
    )
    _FakeToastWidget.instances.clear()
    toast.show_toast(None, "hello")
    created = _FakeToastWidget.instances[-1]
    assert created.moves[-1] == (TOAST_MARGIN, TOAST_MARGIN)
