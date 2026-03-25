"""Workbook secret policy and secure storage helpers."""

from __future__ import annotations

import base64
import ctypes
import os
from pathlib import Path
from typing import Any

from common.constants import APP_NAME
from common.exceptions import ConfigurationError
from common.utils import app_secrets_dir

# One unified password for all sheets (template + reports).
# Kept non-fatal at import time to avoid breaking tooling/bootstrap flows.
# Runtime signing/protection paths should enforce this policy explicitly.
_WORKBOOK_SECRET_STORE_FILENAME = ".workbook_secret.bin"
_WORKBOOK_SECRET_XOR_KEY = 73
_WORKBOOK_SECRET_OBFUSCATED = (
    15, 38, 42, 60, 58, 58, 36, 45, 59, 38, 100, 36, 49, 39, 30, 38, 42, 34, 59, 104, 111, 127
)
_workbook_password_cache: str | None = None
_KEYRING_SERVICE_NAME = f"{APP_NAME}.workbook"
_KEYRING_ACCOUNT_NAME = "workbook_secret"
_WORKBOOK_SECRET_POSIX_USE_KEYRING_ENV_VAR = "FOCUS_WORKBOOK_SECRET_USE_KEYRING"

WORKBOOK_SIGNATURE_VERSION_ENV_VAR = "FOCUS_WORKBOOK_SIGNATURE_VERSION"
WORKBOOK_SIGNATURE_VERSION = os.getenv(WORKBOOK_SIGNATURE_VERSION_ENV_VAR, "v1").strip().lower() or "v1"
if WORKBOOK_SIGNATURE_VERSION not in {"v1"}:
    raise ConfigurationError(
        f"{WORKBOOK_SIGNATURE_VERSION_ENV_VAR} must be one of: v1"
    )


def _default_workbook_password() -> str:
    return "".join(chr(value ^ _WORKBOOK_SECRET_XOR_KEY) for value in _WORKBOOK_SECRET_OBFUSCATED)


def _sanitize_workbook_secret(secret: str) -> str:
    # Excel-adjacent tooling can be sensitive to control characters in passwords.
    return "".join(ch for ch in str(secret) if ord(ch) >= 32 and ord(ch) != 127).strip()


def _workbook_secret_store_path() -> Path:
    return app_secrets_dir(APP_NAME) / _WORKBOOK_SECRET_STORE_FILENAME


def _is_posix() -> bool:
    return os.name == "posix"


def _use_posix_keyring() -> bool:
    if not _is_posix():
        return False
    raw = os.getenv(_WORKBOOK_SECRET_POSIX_USE_KEYRING_ENV_VAR, "1").strip().lower()
    return raw not in {"0", "false", "no", "off"}


def _get_keyring_module() -> Any | None:
    try:
        import keyring  # type: ignore[import-not-found]
    except Exception:
        return None
    return keyring


def _read_workbook_password_from_keyring() -> str:
    if not _use_posix_keyring():
        return ""
    keyring = _get_keyring_module()
    if keyring is None:
        return ""
    try:
        secret = keyring.get_password(_KEYRING_SERVICE_NAME, _KEYRING_ACCOUNT_NAME)
    except Exception:
        return ""
    return str(secret).strip() if secret else ""


def _write_workbook_password_to_keyring(secret: str) -> bool:
    if not _use_posix_keyring():
        return False
    keyring = _get_keyring_module()
    if keyring is None:
        return False
    keyring.set_password(_KEYRING_SERVICE_NAME, _KEYRING_ACCOUNT_NAME, secret)
    return True


def _protect_secret_bytes(secret: bytes) -> bytes:
    if os.name != "nt":
        return base64.b64encode(secret)

    class _DataBlob(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    input_buffer = ctypes.create_string_buffer(secret, len(secret))
    input_blob = _DataBlob(len(secret), ctypes.cast(input_buffer, ctypes.POINTER(ctypes.c_char)))
    output_blob = _DataBlob()

    # Machine-scoped protection keeps the secret readable across local users.
    cryptprotect_local_machine = 0x4
    ok = crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        cryptprotect_local_machine,
        ctypes.byref(output_blob),
    )
    if not ok:
        raise OSError("CryptProtectData failed while storing workbook secret.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(output_blob.pbData)


def _unprotect_secret_bytes(secret: bytes) -> bytes:
    if os.name != "nt":
        return base64.b64decode(secret)

    class _DataBlob(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    input_buffer = ctypes.create_string_buffer(secret, len(secret))
    input_blob = _DataBlob(len(secret), ctypes.cast(input_buffer, ctypes.POINTER(ctypes.c_char)))
    output_blob = _DataBlob()

    ok = crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    )
    if not ok:
        raise OSError("CryptUnprotectData failed while loading workbook secret.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(output_blob.pbData)


def _read_workbook_password_from_store() -> str:
    store_path = _workbook_secret_store_path()
    if not store_path.exists():
        return ""
    protected = store_path.read_bytes()
    if not protected:
        return ""
    try:
        raw = _unprotect_secret_bytes(protected)
    except Exception as exc:
        raise ConfigurationError("Unable to decrypt workbook secret store.") from exc
    try:
        return raw.decode("utf-8").strip()
    except UnicodeDecodeError as exc:
        raise ConfigurationError("Unable to decrypt workbook secret store.") from exc


def _write_workbook_password_to_store(secret: str) -> None:
    protected = _protect_secret_bytes(secret.encode("utf-8"))
    store_path = _workbook_secret_store_path()
    store_path.parent.mkdir(parents=True, exist_ok=True)
    store_path.write_bytes(protected)
    if _is_posix():
        try:
            os.chmod(store_path, 0o600)
        except OSError:
            pass


def get_workbook_password() -> str:
    global _workbook_password_cache
    if _workbook_password_cache is not None:
        return _workbook_password_cache

    stored_secret = _read_workbook_password_from_keyring()
    if stored_secret:
        sanitized_stored_secret = _sanitize_workbook_secret(stored_secret)
        if sanitized_stored_secret != stored_secret and sanitized_stored_secret:
            try:
                _write_workbook_password_to_keyring(sanitized_stored_secret)
            except OSError:
                pass
        _workbook_password_cache = sanitized_stored_secret
        return _workbook_password_cache

    try:
        stored_secret = _read_workbook_password_from_store()
    except ConfigurationError:
        stored_secret = ""
    if stored_secret:
        sanitized_stored_secret = _sanitize_workbook_secret(stored_secret)
        if sanitized_stored_secret != stored_secret and sanitized_stored_secret:
            try:
                _write_workbook_password_to_store(sanitized_stored_secret)
            except OSError:
                pass
        _workbook_password_cache = sanitized_stored_secret
        return _workbook_password_cache

    bootstrap_secret = _sanitize_workbook_secret(_default_workbook_password())
    if not bootstrap_secret:
        _workbook_password_cache = ""
        return _workbook_password_cache
    try:
        try:
            _write_workbook_password_to_keyring(bootstrap_secret)
        except OSError:
            _write_workbook_password_to_store(bootstrap_secret)
        else:
            _write_workbook_password_to_store(bootstrap_secret)
    except OSError:
        # Allow in-memory bootstrap secret when profile storage is unavailable.
        pass
    _workbook_password_cache = bootstrap_secret
    return _workbook_password_cache


def ensure_workbook_secret_policy() -> None:
    if not get_workbook_password():
        raise ConfigurationError("Workbook secret is required and must not be empty.")
