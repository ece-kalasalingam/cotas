"""Compile Qt TS catalogs into QM files (Qt-native i18n workflow)."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _lrelease_executable() -> Path:
    env_root = Path(sys.executable).resolve().parent
    direct_qt_tool = env_root / "Lib" / "site-packages" / "PySide6" / "lrelease.exe"
    if direct_qt_tool.exists():
        return direct_qt_tool

    scripts_dir = env_root / "Scripts"
    wrapper_exe = scripts_dir / "pyside6-lrelease.exe"
    if wrapper_exe.exists():
        return wrapper_exe
    return scripts_dir / "pyside6-lrelease"


def _compile_qm(ts_path: Path, qm_path: Path) -> None:
    executable = str(_lrelease_executable())
    argv = [executable, str(ts_path), "-qm", str(qm_path)]
    exit_code = os.spawnv(os.P_WAIT, executable, argv)  # nosec B606 - trusted local Qt tool invocation
    if exit_code != 0:
        raise RuntimeError(f"lrelease failed with exit code {exit_code}: {' '.join(argv)}")


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    i18n_dir = root / "common" / "i18n"

    ts_files = sorted(i18n_dir.glob("obe_*.ts"))
    if not ts_files:
        print(f"No TS files found under: {i18n_dir}")
        return 1

    for ts_path in ts_files:
        qm_path = ts_path.with_suffix(".qm")
        _compile_qm(ts_path, qm_path)

    print(f"Compiled {len(ts_files)} TS file(s) into QM catalogs in: {i18n_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
