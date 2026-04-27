"""Run local quality checks used by CI and pre-release validation."""

from __future__ import annotations

import argparse
import os
import subprocess  # nosec B404 - trusted local command orchestration
import sys
from pathlib import Path

_PIP_AUDIT_IGNORED_VULNS: tuple[str, ...] = (
    # No fix published as of 2026-04-27; tracked in upstream advisory.
    "GHSA-58qw-9mgm-455v",
)


def _run(
    command: list[str],
    *,
    repo_root: Path,
    env_overrides: dict[str, str] | None = None,
) -> int:
    """Run.
    
    Args:
        command: Parameter value (list[str]).
        repo_root: Parameter value (Path).
    
    Returns:
        int: Return value.
    
    Raises:
        None.
    """
    env = os.environ.copy()
    if env_overrides:
        env.update(env_overrides)
    process = subprocess.run(command, cwd=repo_root, check=False, env=env)  # nosec B603
    return process.returncode


def _parse_args() -> argparse.Namespace:
    """Parse args.
    
    Args:
        None.
    
    Returns:
        argparse.Namespace: Return value.
    
    Raises:
        None.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        choices=("strict", "fast"),
        default="strict",
        help="strict includes dependency audit; fast skips dependency audit for daily iteration.",
    )
    return parser.parse_args()


def main() -> int:
    """Main.
    
    Args:
        None.
    
    Returns:
        int: Return value.
    
    Raises:
        None.
    """
    args = _parse_args()
    repo_root = Path(__file__).resolve().parents[1]
    pip_audit_cache = repo_root / ".pip_audit_cache"
    commands: list[tuple[list[str], dict[str, str] | None]] = [
        ([sys.executable, "-m", "pyflakes", "."], None),
        (
            [
                sys.executable,
                "-m",
                "bandit",
                "-q",
                "-c",
                ".bandit.yaml",
                "-r",
                "common",
                "modules",
                "services",
            ],
            None,
        ),
        ([sys.executable, "scripts/check_ui_strings.py"], None),
        ([sys.executable, "-m", "pytest", "-q", "tests"], None),
    ]
    if args.mode == "strict":
        pip_audit_cache.mkdir(parents=True, exist_ok=True)
        pip_audit_command = [
            sys.executable,
            "-m",
            "pip_audit",
            "--cache-dir",
            str(pip_audit_cache),
        ]
        for vuln_id in _PIP_AUDIT_IGNORED_VULNS:
            pip_audit_command.extend(["--ignore-vuln", vuln_id])
        commands.insert(
            0,
            (pip_audit_command, None),
        )
        commands[-1] = (
            [sys.executable, "-m", "pytest", "-q", "tests"],
            {"RUN_PERF_TESTS": "1"},
        )
    for command, env_overrides in commands:
        code = _run(command, repo_root=repo_root, env_overrides=env_overrides)
        if code != 0:
            return code
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
