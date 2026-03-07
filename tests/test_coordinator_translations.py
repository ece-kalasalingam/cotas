from __future__ import annotations

import re
from pathlib import Path

from common.texts.en import TEXTS as EN_TEXTS
from common.texts.ta_in import TEXTS as TA_TEXTS


def _coordinator_keys_used_in_module() -> set[str]:
    source = Path("modules/coordinator_module.py").read_text(encoding="utf-8")
    return set(re.findall(r't\("((?:coordinator)\.[^"]+)"', source))


def _placeholders(value: str) -> set[str]:
    return set(re.findall(r"\{[^{}]+\}", value))


def test_coordinator_module_keys_exist_in_both_catalogs() -> None:
    keys = _coordinator_keys_used_in_module()
    assert keys
    assert all(key in EN_TEXTS for key in keys)
    assert all(key in TA_TEXTS for key in keys)


def test_coordinator_tamil_placeholders_match_english() -> None:
    for key in sorted(_coordinator_keys_used_in_module()):
        assert _placeholders(TA_TEXTS[key]) == _placeholders(EN_TEXTS[key])


def test_coordinator_tamil_strings_are_not_english_fallbacks() -> None:
    for key in sorted(_coordinator_keys_used_in_module()):
        assert TA_TEXTS[key] != EN_TEXTS[key]
