"""Run local quality checks used by CI and pre-release validation."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _run(command: list[str], *, repo_root: Path) -> int:
    process = subprocess.run(command, cwd=repo_root, check=False)
    return process.returncode


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    commands = [
        [sys.executable, "-m", "pyflakes", "."],
        [sys.executable, "scripts/check_ui_strings.py"],
        [sys.executable, "-m", "pytest", "-q"],
    ]
    for command in commands:
        code = _run(command, repo_root=repo_root)
        if code != 0:
            return code
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
