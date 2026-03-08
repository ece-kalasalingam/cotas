from __future__ import annotations

import ast
import re
from pathlib import Path

from common.texts.en import TEXTS as EN_TEXTS
from common.texts.ta_in import TEXTS as TA_TEXTS
from modules import __file__ as MODULES_INIT_FILE


def _normalized_text(value: str) -> str:
    return " ".join(value.split()).strip(".,;:!?")


def _coordinator_keys_used_in_module() -> set[str]:
    module_path = Path(MODULES_INIT_FILE).resolve().parent / "coordinator_module.py"
    source = module_path.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(module_path))
    keys: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "t":
            continue
        if not node.args:
            continue
        first_arg = node.args[0]
        if not isinstance(first_arg, ast.Constant) or not isinstance(first_arg.value, str):
            continue
        if first_arg.value.startswith("coordinator."):
            keys.add(first_arg.value)
    return keys


def _placeholders(value: str) -> set[str]:
    return set(re.findall(r"\{[^{}]+\}", value))


def test_coordinator_module_keys_exist_in_both_catalogs() -> None:
    keys = _coordinator_keys_used_in_module()
    assert keys, (
        "No coordinator translation keys were found in coordinator_module.py. "
        "This likely means the key-extraction logic or module layout has changed."
    )
    missing_en = keys - EN_TEXTS.keys()
    assert not missing_en, f"Missing English coordinator keys: {sorted(missing_en)}"
    missing_ta = keys - TA_TEXTS.keys()
    assert not missing_ta, f"Missing Tamil coordinator keys: {sorted(missing_ta)}"


def test_coordinator_tamil_placeholders_match_english() -> None:
    for key in sorted(_coordinator_keys_used_in_module()):
        ta_placeholders = _placeholders(TA_TEXTS[key])
        en_placeholders = _placeholders(EN_TEXTS[key])
        assert ta_placeholders == en_placeholders, (
            f"Placeholder mismatch for key {key}: "
            f"TA={sorted(ta_placeholders)}, EN={sorted(en_placeholders)}"
        )


def test_coordinator_tamil_strings_are_not_english_fallbacks() -> None:
    for key in sorted(_coordinator_keys_used_in_module()):
        en_value = EN_TEXTS[key]
        ta_value = TA_TEXTS[key]
        # Allow short labels/acronyms/proper nouns to remain identical.
        if len(en_value) <= 10:
            continue
        assert _normalized_text(ta_value) != _normalized_text(en_value), (
            f"Possible English fallback for key {key}: "
            f"TA={ta_value!r}, EN={en_value!r}"
        )
