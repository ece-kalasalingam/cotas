from __future__ import annotations

import ast
from pathlib import Path

from modules import coordinator_module as coordinator


def _ns_keys_used_in_file(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    keys: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Subscript):
            continue
        if not isinstance(node.value, ast.Name) or node.value.id != "ns":
            continue
        key_node = node.slice
        if isinstance(key_node, ast.Constant) and isinstance(key_node.value, str):
            keys.add(key_node.value)
    return keys


def _coordinator_helper_ns_key_usage() -> dict[str, set[str]]:
    helper_root = Path("modules") / "coordinator"
    usage: dict[str, set[str]] = {}
    for path in sorted(helper_root.rglob("*.py")):
        if path.name == "__init__.py":
            continue
        for key in _ns_keys_used_in_file(path):
            usage.setdefault(key, set()).add(str(path).replace("\\", "/"))
    return usage


def test_coordinator_helper_ns_keys_exist_in_coordinator_module_globals() -> None:
    usage = _coordinator_helper_ns_key_usage()
    assert usage, "No coordinator helper ns[...] key usage found; test assumptions may be outdated."

    available = set(vars(coordinator).keys())
    missing = sorted(key for key in usage if key not in available)

    assert not missing, (
        "Coordinator helper modules reference ns keys that are missing from "
        "modules.coordinator_module globals(): "
        + ", ".join(f"{key} (used in {sorted(usage[key])})" for key in missing)
    )
