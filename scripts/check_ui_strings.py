"""Fail when user-facing UI strings are hardcoded in UI modules."""

from __future__ import annotations

import ast
import sys
from pathlib import Path

CALL_ARG_INDEXES: dict[str, tuple[int, ...]] = {
    "QLabel": (0,),
    "QPushButton": (0,),
    "QGroupBox": (0,),
    "QToolBar": (0,),
    "QAction": (1,),
    "showMessage": (0,),
    "setText": (0,),
    "setWindowTitle": (0,),
    "addAction": (0,),
    "warning": (1, 2),
    "information": (1, 2),
    "critical": (1, 2),
    "getSaveFileName": (1, 2, 3),
    "getOpenFileName": (1, 2, 3),
}


def _call_name(node: ast.Call) -> str | None:
    func = node.func
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return func.attr
    return None


def _is_string_literal(node: ast.AST) -> bool:
    return isinstance(node, ast.Constant) and isinstance(node.value, str)


def _is_fstring(node: ast.AST) -> bool:
    return isinstance(node, ast.JoinedStr)


def _find_violations(source: str, file_path: Path) -> list[str]:
    tree = ast.parse(source, filename=str(file_path))
    violations: list[str] = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue

        name = _call_name(node)
        if name not in CALL_ARG_INDEXES:
            continue

        for idx in CALL_ARG_INDEXES[name]:
            if idx >= len(node.args):
                continue
            arg = node.args[idx]
            if _is_string_literal(arg) or _is_fstring(arg):
                violations.append(
                    f"{file_path}:{arg.lineno}:{arg.col_offset + 1} "
                    f"hardcoded string in `{name}`"
                )
    return violations


def _target_files(repo_root: Path) -> list[Path]:
    files = [repo_root / "main.py", repo_root / "main_window.py"]
    files.extend((repo_root / "modules").glob("*.py"))
    return [f for f in files if f.is_file()]


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    failures: list[str] = []
    for file_path in _target_files(repo_root):
        source = file_path.read_text(encoding="utf-8")
        failures.extend(_find_violations(source, file_path))

    if failures:
        print("UI hardcoded-string check failed:")
        for failure in failures:
            print(f"  - {failure}")
        print("Use common.texts.t(...) keys instead of inline literals.")
        return 1

    print("UI hardcoded-string check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
