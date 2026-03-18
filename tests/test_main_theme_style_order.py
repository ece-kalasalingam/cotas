from __future__ import annotations

import importlib

import pytest


def _reloaded_main():
    import main as main_mod

    return importlib.reload(main_mod)


def test_fusion_style_runs_before_theme_setup(monkeypatch: pytest.MonkeyPatch) -> None:
    main_mod = _reloaded_main()
    events: list[str] = []

    class _FakeApp:
        def setStyle(self, style: str) -> None:  # noqa: N802
            events.append(f"style:{style}")

        def setOrganizationName(self, _name: str) -> None:  # noqa: N802
            return None

        def setApplicationName(self, _name: str) -> None:  # noqa: N802
            return None

        def setApplicationDisplayName(self, _name: str) -> None:  # noqa: N802
            return None

    fake_app = _FakeApp()

    class _FakeQApplication:
        @staticmethod
        def setHighDpiScaleFactorRoundingPolicy(_policy) -> None:  # noqa: N802
            return None

        def __new__(cls, *_args, **_kwargs):
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
