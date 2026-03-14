"""Run local quality checks used by CI and pre-release validation."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _run(command: list[str], *, repo_root: Path) -> int:
    process = subprocess.run(command, cwd=repo_root, check=False)
    return process.returncode


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        choices=("strict", "fast"),
        default="strict",
        help="strict includes dependency audit; fast skips dependency audit for daily iteration.",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    repo_root = Path(__file__).resolve().parents[1]
    commands: list[list[str]] = [
        [sys.executable, "-m", "pyflakes", "."],
        [sys.executable, "-m", "bandit", "-q", "-c", ".bandit.yaml", "-r", "common", "modules", "services"],
        [sys.executable, "scripts/check_ui_strings.py"],
        [sys.executable, "-m", "pytest", "-q"],
    ]
    if args.mode == "strict":
        commands.insert(0, [sys.executable, "-m", "pip_audit"])
    for command in commands:
        code = _run(command, repo_root=repo_root)
        if code != 0:
            return code
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
