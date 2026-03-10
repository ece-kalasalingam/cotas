from __future__ import annotations

import ast
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INTERNAL_ROOTS = ("common", "domain", "services", "modules")
IGNORED_PARTS = {"__pycache__", "tests"}


def _python_files_under(root: Path) -> list[Path]:
    files: list[Path] = []
    if not root.exists():
        return files
    for path in root.rglob("*.py"):
        parts = set(path.as_posix().split("/"))
        if parts & IGNORED_PARTS:
            continue
        files.append(path)
    return files


def _internal_import_roots(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    roots: set[str] = set()
    for node in ast.walk(tree):
        module_name = None
        if isinstance(node, ast.Import):
            for alias in node.names:
                module_name = alias.name
                if module_name:
                    root = module_name.split(".", 1)[0]
                    if root in INTERNAL_ROOTS:
                        roots.add(root)
        elif isinstance(node, ast.ImportFrom):
            module_name = node.module or ""
            if module_name:
                root = module_name.split(".", 1)[0]
                if root in INTERNAL_ROOTS:
                    roots.add(root)
    return roots


def _violations_for_layer(layer: str, forbidden_roots: set[str]) -> list[tuple[str, list[str]]]:
    violations: list[tuple[str, list[str]]] = []
    for file in _python_files_under(REPO_ROOT / layer):
        imported = _internal_import_roots(file)
        bad = sorted(imported & forbidden_roots)
        if bad:
            rel = file.relative_to(REPO_ROOT).as_posix()
            violations.append((rel, bad))
    return violations


def test_common_layer_imports_no_higher_layers() -> None:
    violations = _violations_for_layer("common", {"domain", "services", "modules"})
    assert not violations, f"common layer import violations: {violations}"


def test_domain_layer_imports_no_ui_or_service_layers() -> None:
    violations = _violations_for_layer("domain", {"services", "modules"})
    assert not violations, f"domain layer import violations: {violations}"


def test_services_layer_imports_no_ui_layer() -> None:
    violations = _violations_for_layer("services", {"modules"})
    assert not violations, f"services layer import violations: {violations}"
