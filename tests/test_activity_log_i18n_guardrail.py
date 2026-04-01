from __future__ import annotations

import ast
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MODULES_ROOT = REPO_ROOT / "modules"
TARGET_MODULES = tuple(
    path
    for path in MODULES_ROOT.rglob("*.py")
    if "__pycache__" not in path.parts and path.is_file()
)


def _literal_channels(node: ast.Call) -> set[str] | None:
    """Literal channels.
    
    Args:
        node: Parameter value (ast.Call).
    
    Returns:
        set[str] | None: Return value.
    
    Raises:
        None.
    """
    for kw in node.keywords:
        if kw.arg != "channels":
            continue
        value = kw.value
        if isinstance(value, (ast.Tuple, ast.List)):
            items: set[str] = set()
            for elt in value.elts:
                if isinstance(elt, ast.Constant) and isinstance(elt.value, str):
                    items.add(elt.value)
            return items
    return None


def _violations_for_file(path: Path) -> list[str]:
    """Violations for file.
    
    Args:
        path: Parameter value (Path).
    
    Returns:
        list[str]: Return value.
    
    Raises:
        None.
    """
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    violations: list[str] = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        fn = node.func
        if not isinstance(fn, ast.Attribute):
            continue
        if not isinstance(fn.value, ast.Attribute):
            continue
        if not isinstance(fn.value.value, ast.Name):
            continue
        if fn.value.value.id != "self" or fn.value.attr != "_runtime":
            continue

        if fn.attr in {"append_user_log", "publish_status"}:
            violations.append(f"{path.name}:{node.lineno} uses self._runtime.{fn.attr}(...)")
            continue

        if fn.attr == "notify_message":
            channels = _literal_channels(node)
            if channels is None:
                # notify_message defaults to ("status", "activity_log"), so this is disallowed.
                violations.append(
                    f"{path.name}:{node.lineno} uses self._runtime.notify_message(...) without channels"
                )
                continue
            if "status" in channels or "activity_log" in channels:
                violations.append(
                    f"{path.name}:{node.lineno} uses self._runtime.notify_message(...) for status/activity_log"
                )
    return violations


def test_all_modules_activity_log_messages_are_key_based() -> None:
    """Test all modules activity log messages are key based.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    violations: list[str] = []
    for path in TARGET_MODULES:
        violations.extend(_violations_for_file(path))
    assert not violations, "Activity-log i18n guardrail violations:\n" + "\n".join(violations)
