from __future__ import annotations

import importlib
from pathlib import Path

import pytest

from common.exceptions import ConfigurationError


def _reloaded_workbook_secret():
    """Reloaded workbook secret.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    import common.workbook_integrity.workbook_secret as workbook_secret_mod

    return importlib.reload(workbook_secret_mod)


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


def test_secret_policy_auto_provisions_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Test secret policy auto provisions secret.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "wb_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)

    secret = workbook_secret_mod.get_workbook_password()

    if not (isinstance(secret, str)):
        raise AssertionError('assertion failed')
    if not (secret):
        raise AssertionError('assertion failed')
    if not (store_path.exists()):
        raise AssertionError('assertion failed')
    if not (store_path.read_bytes()):
        raise AssertionError('assertion failed')
    workbook_secret_mod.ensure_workbook_secret_policy()


def test_startup_validation_accepts_auto_secret(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Test startup validation accepts auto secret.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "auto_startup_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    main_mod = _reloaded_main()
    monkeypatch.setattr(main_mod, "ensure_workbook_secret_policy", workbook_secret_mod.ensure_workbook_secret_policy)

    if main_mod._validate_startup_workbook_password(None) is not None:
        raise AssertionError('assertion failed')


def test_secret_policy_recovers_from_unreadable_store(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Test secret policy recovers from unreadable store.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "broken_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    store_path.write_bytes(b"not-a-valid-secret-store")

    secret = workbook_secret_mod.get_workbook_password()

    if not (secret):
        raise AssertionError('assertion failed')
    if not (isinstance(secret, str)):
        raise AssertionError('assertion failed')


def test_secret_policy_sanitizes_control_chars_in_stored_secret(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Test secret policy sanitizes control chars in stored secret.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "legacy_secret.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)
    workbook_secret_mod._write_workbook_password_to_store("abc\x7fdef")  # type: ignore[attr-defined]

    secret = workbook_secret_mod.get_workbook_password()

    expected_secret = "abc" + "def"
    if not (secret == expected_secret):
        raise AssertionError('assertion failed')


def test_signature_version_env_validation_on_import(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test signature version env validation on import.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setenv("FOCUS_WORKBOOK_SIGNATURE_VERSION", "v2")
    with pytest.raises(ConfigurationError):
        _reloaded_workbook_secret()


def test_protect_unprotect_non_windows_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test protect unprotect non windows paths.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod.os, "name", "posix", raising=False)
    protected = workbook_secret_mod._protect_secret_bytes(b"secret")
    if not (isinstance(protected, bytes)):
        raise AssertionError('assertion failed')
    if not (workbook_secret_mod._unprotect_secret_bytes(protected) == b"secret"):
        raise AssertionError('assertion failed')


def test_protect_windows_failure_branch(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test protect windows failure branch.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod.os, "name", "nt", raising=False)

    class _Crypt32:
        @staticmethod
        def CryptProtectData(*_args, **_kwargs) -> int:
            """Cryptprotectdata.
            
            Args:
                _args: Parameter value.
                _kwargs: Parameter value.
            
            Returns:
                int: Return value.
            
            Raises:
                None.
            """
            return 0

    class _Kernel32:
        @staticmethod
        def LocalFree(*_args, **_kwargs) -> int:
            """Localfree.
            
            Args:
                _args: Parameter value.
                _kwargs: Parameter value.
            
            Returns:
                int: Return value.
            
            Raises:
                None.
            """
            return 0

    class _Windll:
        crypt32 = _Crypt32()
        kernel32 = _Kernel32()

    monkeypatch.setattr(workbook_secret_mod.ctypes, "windll", _Windll(), raising=False)
    with pytest.raises(OSError):
        workbook_secret_mod._protect_secret_bytes(b"x")


def test_read_store_empty_and_decode_error_paths(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Test read store empty and decode error paths.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    store_path = tmp_path / "store.bin"
    monkeypatch.setattr(workbook_secret_mod, "_workbook_secret_store_path", lambda: store_path)

    store_path.write_bytes(b"")
    if not (workbook_secret_mod._read_workbook_password_from_store() == ""):
        raise AssertionError('assertion failed')

    store_path.write_bytes(b"abc")
    monkeypatch.setattr(workbook_secret_mod, "_unprotect_secret_bytes", lambda _b: b"\xff")
    with pytest.raises(ConfigurationError):
        workbook_secret_mod._read_workbook_password_from_store()


def test_get_password_fallback_and_empty_secret_branches(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test get password fallback and empty secret branches.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_keyring", lambda: "")
    monkeypatch.setattr(workbook_secret_mod, "_write_workbook_password_to_keyring", lambda _s: False)

    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "abc\x7fdef")
    monkeypatch.setattr(workbook_secret_mod, "_write_workbook_password_to_store", lambda _s: (_ for _ in ()).throw(OSError("locked")))
    expected_secret = "abc" + "def"
    if not (workbook_secret_mod.get_workbook_password() == expected_secret):
        raise AssertionError('assertion failed')

    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "")
    monkeypatch.setattr(workbook_secret_mod, "_default_workbook_password", lambda: "\x7f\t")
    if not (workbook_secret_mod.get_workbook_password() == ""):
        raise AssertionError('assertion failed')

    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_default_workbook_password", lambda: "abc")
    monkeypatch.setattr(workbook_secret_mod, "_write_workbook_password_to_store", lambda _s: (_ for _ in ()).throw(OSError("io")))
    if not (workbook_secret_mod.get_workbook_password() == "abc"):
        raise AssertionError('assertion failed')


def test_get_password_uses_posix_keyring_when_available(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test get password uses posix keyring when available.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_keyring", lambda: "from-keyring")
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "from-store")

    if not (workbook_secret_mod.get_workbook_password() == "from-keyring"):
        raise AssertionError('assertion failed')


def test_get_password_bootstrap_writes_to_keyring_and_store(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test get password bootstrap writes to keyring and store.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "_workbook_password_cache", None, raising=False)
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_keyring", lambda: "")
    monkeypatch.setattr(workbook_secret_mod, "_read_workbook_password_from_store", lambda: "")
    monkeypatch.setattr(workbook_secret_mod, "_default_workbook_password", lambda: "abc")
    writes: dict[str, list[str]] = {"keyring": [], "store": []}
    monkeypatch.setattr(
        workbook_secret_mod,
        "_write_workbook_password_to_keyring",
        lambda s: writes["keyring"].append(s) or True,
    )
    monkeypatch.setattr(
        workbook_secret_mod,
        "_write_workbook_password_to_store",
        lambda s: writes["store"].append(s),
    )

    if not (workbook_secret_mod.get_workbook_password() == "abc"):
        raise AssertionError('assertion failed')
    if not (writes["keyring"] == ["abc"]):
        raise AssertionError('assertion failed')
    if not (writes["store"] == ["abc"]):
        raise AssertionError('assertion failed')


def test_ensure_workbook_secret_policy_raises_when_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test ensure workbook secret policy raises when empty.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_secret_mod = _reloaded_workbook_secret()
    monkeypatch.setattr(workbook_secret_mod, "get_workbook_password", lambda: "")
    with pytest.raises(ConfigurationError):
        workbook_secret_mod.ensure_workbook_secret_policy()

