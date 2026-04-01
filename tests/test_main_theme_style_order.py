from __future__ import annotations

import importlib

import pytest


def _reloaded_main():
    """Reloaded main.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    import main as main_mod

    return importlib.reload(main_mod)


def test_fusion_style_runs_before_theme_setup(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test fusion style runs before theme setup.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    main_mod = _reloaded_main()
    events: list[str] = []

    class _FakeApp:
        def setStyle(self, style: str) -> None:  # noqa: N802
            """Setstyle.
            
            Args:
                style: Parameter value (str).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            events.append(f"style:{style}")

        def setOrganizationName(self, _name: str) -> None:  # noqa: N802
            """Setorganizationname.
            
            Args:
                _name: Parameter value (str).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

        def setApplicationName(self, _name: str) -> None:  # noqa: N802
            """Setapplicationname.
            
            Args:
                _name: Parameter value (str).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

        def setApplicationDisplayName(self, _name: str) -> None:  # noqa: N802
            """Setapplicationdisplayname.
            
            Args:
                _name: Parameter value (str).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

    fake_app = _FakeApp()

    class _FakeQApplication:
        @staticmethod
        def setHighDpiScaleFactorRoundingPolicy(_policy) -> None:  # noqa: N802
            """Sethighdpiscalefactorroundingpolicy.
            
            Args:
                _policy: Parameter value.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

        def __new__(cls, *_args, **_kwargs):
            """New.
            
            Args:
                _args: Parameter value.
                _kwargs: Parameter value.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return fake_app

    monkeypatch.setattr(main_mod, "QApplication", _FakeQApplication)
    monkeypatch.setattr(main_mod, "_setup_system_theme", lambda: events.append("theme"))
    monkeypatch.setattr(main_mod, "validate_blueprint_registry_contracts", lambda: None)
    monkeypatch.setattr(main_mod, "configure_app_logging", lambda _app_name: None)
    monkeypatch.setattr(main_mod, "get_ui_language_preference", lambda _app_name: "en")
    monkeypatch.setattr(main_mod, "set_language", lambda _code: None)
    monkeypatch.setattr(main_mod, "_install_excepthook", lambda: None)
    monkeypatch.setattr(main_mod, "_validate_startup_workbook_password", lambda _app: 1)

    result = main_mod.main()

    assert result == 1
    assert events == ["style:Fusion", "theme"]
