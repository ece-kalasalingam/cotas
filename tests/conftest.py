from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from uuid import uuid4

import pytest


WINDOWS_ACL_COMPAT_MARK = "windows_acl_compat"
WINDOWS_ACL_COMPAT_ENV = "FOCUS_TEST_ENABLE_WINDOWS_ACL_COMPAT"


def _pytest_tmp_env_root() -> Path:
    root = Path.cwd() / ".pytest_tmp_env"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _set_fresh_pytest_temp_root() -> None:
    """Pin pytest temp root inside .pytest_tmp_env to avoid root-level clutter."""
    if not sys.platform.startswith("win"):
        return

    if os.environ.get("PYTEST_DEBUG_TEMPROOT"):
        return

    root = _pytest_tmp_env_root() / "temproot"
    root.mkdir(parents=True, exist_ok=True)
    os.environ["PYTEST_DEBUG_TEMPROOT"] = str(root)


def _set_fresh_process_temp_root() -> None:
    """Force stdlib tempfile to a stable writable dir under .pytest_tmp_env."""
    if not sys.platform.startswith("win"):
        return

    root = _pytest_tmp_env_root() / "runtime_temp"
    root.mkdir(parents=True, exist_ok=True)
    os.environ["TEMP"] = str(root)
    os.environ["TMP"] = str(root)
    tempfile.tempdir = str(root)


def _apply_windows_pytest_tempdir_compat() -> None:
    """Avoid WinError 5 on sandboxed Windows runs with restrictive tempdir ACLs."""
    if not sys.platform.startswith("win"):
        return

    try:
        import _pytest.pathlib as pytest_pathlib
    except Exception:
        return

    original_make_numbered_dir = pytest_pathlib.make_numbered_dir

    def _patched_make_numbered_dir(root, prefix, mode=0o700):  # type: ignore[no-untyped-def]
        # On this runtime, mode=0o700 can produce inaccessible directories.
        return original_make_numbered_dir(root, prefix, mode=0o777)

    pytest_pathlib.make_numbered_dir = _patched_make_numbered_dir


def _is_pytest_temp_path(path_value: object) -> bool:
    text = str(path_value).replace("/", "\\").lower()
    return "pytest-of-" in text or "\\.pytest_" in text or "\\pytest-" in text


def _patch_windows_pytest_temp_mkdir_mode() -> None:
    """Keep pytest temp paths writable on Windows sandboxed ACL setups."""
    if not sys.platform.startswith("win"):
        return

    original_mkdir = Path.mkdir

    def _patched_mkdir(self, mode=0o777, parents=False, exist_ok=False):  # type: ignore[no-untyped-def]
        if mode == 0o700 and _is_pytest_temp_path(self):
            mode = 0o777
        return original_mkdir(self, mode=mode, parents=parents, exist_ok=exist_ok)

    Path.mkdir = _patched_mkdir


def _patch_windows_tempfile_mkdtemp() -> None:
    """Create TemporaryDirectory paths with writable ACL behavior on Windows sandbox."""
    if not sys.platform.startswith("win"):
        return

    def _patched_mkdtemp(suffix=None, prefix=None, dir=None):  # type: ignore[no-untyped-def]
        base = Path(dir or tempfile.gettempdir())
        base.mkdir(parents=True, exist_ok=True)
        file_prefix = prefix or "tmp"
        file_suffix = suffix or ""
        for _ in range(1000):
            candidate = base / f"{file_prefix}{uuid4().hex}{file_suffix}"
            try:
                candidate.mkdir(mode=0o777)
                return str(candidate)
            except FileExistsError:
                continue
        raise FileExistsError("No usable temporary directory name found.")

    tempfile.mkdtemp = _patched_mkdtemp  # type: ignore[assignment]


def _patch_windows_pytest_chmod_compat(monkeypatch: pytest.MonkeyPatch) -> None:
    """Suppress known tempdir chmod PermissionError only for explicitly opted-in tests."""
    if not sys.platform.startswith("win"):
        return

    original_chmod = os.chmod

    def _patched_chmod(path, mode, *args, **kwargs):  # type: ignore[no-untyped-def]
        try:
            return original_chmod(path, mode, *args, **kwargs)
        except PermissionError:
            if _is_pytest_temp_path(path):
                return None
            raise

    monkeypatch.setattr(os, "chmod", _patched_chmod)


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        f"{WINDOWS_ACL_COMPAT_MARK}: enable scoped Windows ACL compatibility patches for tempdir operations.",
    )


@pytest.fixture(autouse=True)
def _windows_acl_compat_patches(request: pytest.FixtureRequest, monkeypatch: pytest.MonkeyPatch) -> None:
    if not sys.platform.startswith("win"):
        return
    marker_enabled = request.node.get_closest_marker(WINDOWS_ACL_COMPAT_MARK) is not None
    env_enabled = os.getenv(WINDOWS_ACL_COMPAT_ENV, "").strip() == "1"
    if not (marker_enabled or env_enabled):
        return
    _patch_windows_pytest_chmod_compat(monkeypatch)


_apply_windows_pytest_tempdir_compat()
_set_fresh_process_temp_root()
_set_fresh_pytest_temp_root()
_patch_windows_pytest_temp_mkdir_mode()
_patch_windows_tempfile_mkdtemp()

@pytest.fixture(autouse=True)
def _disable_native_file_dialogs(monkeypatch: pytest.MonkeyPatch) -> None:
    """Prevent native OS file dialogs from appearing during automated tests."""
    try:
        from PySide6.QtWidgets import QFileDialog
    except Exception:
        return

    monkeypatch.setattr(QFileDialog, "getOpenFileName", staticmethod(lambda *_a, **_k: ("", "")), raising=False)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", staticmethod(lambda *_a, **_k: ("", "")), raising=False)
    monkeypatch.setattr(QFileDialog, "getOpenFileNames", staticmethod(lambda *_a, **_k: ([], "")), raising=False)

