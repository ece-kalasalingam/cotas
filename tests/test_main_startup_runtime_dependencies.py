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


def test_validate_startup_runtime_dependencies_blocks_on_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test validate startup runtime dependencies blocks on missing.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    main_mod = _reloaded_main()
    captured: dict[str, str] = {}
    monkeypatch.setattr(
        main_mod,
        "missing_runtime_dependency_packages",
        lambda: ("openpyxl", "python-docx"),
    )
    monkeypatch.setattr(
        main_mod,
        "t",
        lambda key, **kwargs: f"{key}:{kwargs.get('packages', '')}",
    )
    monkeypatch.setattr(
        main_mod,
        "_show_startup_error_dialog",
        lambda *, title, message: captured.update({"title": title, "message": message}),
    )

    result = main_mod._validate_startup_runtime_dependencies(None)

    if not (result == 1):
        raise AssertionError('assertion failed')
    if not (captured["title"] == main_mod.APP_NAME):
        raise AssertionError('assertion failed')
    if "openpyxl, python-docx" not in captured["message"]:
        raise AssertionError('assertion failed')


def test_validate_startup_runtime_dependencies_passes_when_present(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test validate startup runtime dependencies passes when present.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    main_mod = _reloaded_main()
    monkeypatch.setattr(main_mod, "missing_runtime_dependency_packages", lambda: ())
    if main_mod._validate_startup_runtime_dependencies(None) is not None:
        raise AssertionError('assertion failed')
