from __future__ import annotations

import importlib
from pathlib import Path

import pytest
from common.exceptions import ConfigurationError


def _reloaded_workbook_secret():
    import common.workbook_secret as workbook_secret_mod

    return importlib.reload(workbook_secret_mod)


def _reloaded_main():
    import main as main_mod

    return importlib.reload(main_mod)


def test_secret_policy_auto_provisions_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "wb_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)

    secret = workbook_secret_mod.get_workbook_password()

    assert isinstance(secret, str)
    assert secret
    assert store_path.exists()
    assert store_path.read_bytes()
    workbook_secret_mod.ensure_workbook_secret_policy()


def test_startup_validation_accepts_auto_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "auto_startup_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    main_mod = _reloaded_main()
    monkeypatch.setattr(main_mod, "ensure_workbook_secret_policy", workbook_secret_mod.ensure_workbook_secret_policy)

    assert main_mod._validate_startup_workbook_password(None) is None


def test_secret_policy_recovers_from_unreadable_store(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "broken_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    store_path.write_bytes(b"not-a-valid-secret-store")

    secret = workbook_secret_mod.get_workbook_password()

    assert secret
    assert isinstance(secret, str)


def test_secret_policy_sanitizes_control_chars_in_stored_secret(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "legacy_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    workbook_secret_mod._write_workbook_password_to_store("abc\x7fdef")  # type: ignore[attr-defined]

    secret = workbook_secret_mod.get_workbook_password()

    assert secret == "abcdef"


def test_signature_version_env_validation_on_import(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("FOCUS_WORKBOOK_SIGNATURE_VERSION", "v2")
    with pytest.raises(ConfigurationError):
        _reloaded_workbook_secret()


def test_protect_unprotect_non_windows_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod.os, "name", "posix", raising=False)
    protected = workbook_secret_mod._protect_secret_bytes(b"secret")
    assert isinstance(protected, bytes)
    assert workbook_secret_mod._unprotect_secret_bytes(protected) == b"secret"


def test_protect_windows_failure_branch(monkeypatch: pytest.MonkeyPatch) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod.os, "name", "nt", raising=False)

    class _Crypt32:
        @staticmethod
        def CryptProtectData(*_args, **_kwargs) -> int:
            return 0

    class _Kernel32:
        @staticmethod
        def LocalFree(*_args, **_kwargs) -> int:
            return 0

    class _Windll:
        crypt32 = _Crypt32()
        kernel32 = _Kernel32()

    monkeypatch.setattr(workbook_secret_mod.ctypes, "windll", _Windll(), raising=False)
    with pytest.raises(OSError):
        workbook_secret_mod._protect_secret_bytes(b"x")


def test_read_store_empty_and_decode_error_paths(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "store.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)

    store_path.write_bytes(b"")
    assert workbook_secret_mod._read_workbook_password_from_store() == ""

    store_path.write_bytes(b"abc")
    monkeypatch.setattr(workbook_secret_mod, "_unprotect_secret_bytes", lambda _b: b"\xff")
    with pytest.raises(ConfigurationError):
        workbook_secret_mod._read_workbook_password_from_store()


def test_get_password_fallback_and_empty_secret_branches(monkeypatch: pytest.MonkeyPatch) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)

    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "abc\x7fdef")
    monkeypatch.setattr(workbook_secret_mod, "_write_workbook_password_to_store", lambda _s: (_ for _ in ()).throw(OSError("locked")))
    assert workbook_secret_mod.get_workbook_password() == "abcdef"

    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "")
    monkeypatch.setattr(workbook_secret_mod, "_default_workbook_password", lambda: "\x7f\t")
    assert workbook_secret_mod.get_workbook_password() == ""

    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_default_workbook_password", lambda: "abc")
    monkeypatch.setattr(workbook_secret_mod, "_write_workbook_password_to_store", lambda _s: (_ for _ in ()).throw(OSError("io")))
    assert workbook_secret_mod.get_workbook_password() == "abc"


def test_ensure_workbook_secret_policy_raises_when_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "get_workbook_password", lambda: "")
    with pytest.raises(ConfigurationError):
        workbook_secret_mod.ensure_workbook_secret_policy()
