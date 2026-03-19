"""Run local quality checks used by CI and pre-release validation."""

from __future__ import annotations

import argparse
import subprocess  # nosec B404 - trusted local command orchestration
import sys
from pathlib import Path


def _run(command: list[str], *, repo_root: Path) -> int:
    process = subprocess.run(command, cwd=repo_root, check=False)  # nosec B603
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
    pip_audit_cache = repo_root / ".pip_audit_cache"
    commands: list[list[str]] = [
        [sys.executable, "-m", "pyflakes", "."],
        [sys.executable, "-m", "bandit", "-q", "-c", ".bandit.yaml", "-r", "common", "modules", "services"],
        [sys.executable, "scripts/check_ui_strings.py"],
        [sys.executable, "-m", "pytest", "-q", "tests"],
    ]
    if args.mode == "strict":
        pip_audit_cache.mkdir(parents=True, exist_ok=True)
        commands.insert(0, [sys.executable, "-m", "pip_audit", "--cache-dir", str(pip_audit_cache)])
    for command in commands:
        code = _run(command, repo_root=repo_root)
        if code != 0:
            return code
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
