from __future__ import annotations

import importlib
from pathlib import Path

import pytest


def _reloaded_constants():
    import common.constants as constants_mod

    return importlib.reload(constants_mod)


def _reloaded_main():
    import main as main_mod

    return importlib.reload(main_mod)


def test_secret_policy_auto_provisions_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    constants_mod = _reloaded_constants()
    store_path = tmp_path / "wb_secret.bin"
    monkeypatch.setattr(constants_mod, "_workbook_secret_store_path", lambda: store_path)

    secret = constants_mod.get_workbook_password()

    assert isinstance(secret, str)
    assert secret
    assert store_path.exists()
    assert store_path.read_bytes()
    constants_mod.ensure_workbook_secret_policy()


def test_startup_validation_accepts_auto_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    constants_mod = _reloaded_constants()
    store_path = tmp_path / "auto_startup_secret.bin"
    monkeypatch.setattr(constants_mod, "_workbook_secret_store_path", lambda: store_path)
    main_mod = _reloaded_main()
    monkeypatch.setattr(main_mod, "ensure_workbook_secret_policy", constants_mod.ensure_workbook_secret_policy)

    assert main_mod._validate_startup_workbook_password(None) is None


def test_secret_policy_recovers_from_unreadable_store(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    constants_mod = _reloaded_constants()
    store_path = tmp_path / "broken_secret.bin"
    monkeypatch.setattr(constants_mod, "_workbook_secret_store_path", lambda: store_path)
    store_path.write_bytes(b"not-a-valid-secret-store")

    secret = constants_mod.get_workbook_password()

    assert secret
    assert isinstance(secret, str)


def test_secret_policy_sanitizes_control_chars_in_stored_secret(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    constants_mod = _reloaded_constants()
    store_path = tmp_path / "legacy_secret.bin"
    monkeypatch.setattr(constants_mod, "_workbook_secret_store_path", lambda: store_path)
    constants_mod._write_workbook_password_to_store("abc\x7fdef")  # type: ignore[attr-defined]

    secret = constants_mod.get_workbook_password()

    assert secret == "abcdef"
