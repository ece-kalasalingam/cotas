from __future__ import annotations

import ast
import html
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
MODULES_ROOT = REPO_ROOT / "modules"
LOCALES = ("obe_hi_IN.ts", "obe_ta_IN.ts", "obe_te_IN.ts")
KEY_CALLS = {"notify_message_key", "publish_status_key", "_publish_status_key", "build_status_message", "_build_status_message"}


def _collect_module_i18n_keys() -> set[str]:
    """Collect module i18n keys.
    
    Args:
        None.
    
    Returns:
        set[str]: Return value.
    
    Raises:
        None.
    """
    keys: set[str] = set()
    for path in MODULES_ROOT.rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            fn = node.func
            name = None
            if isinstance(fn, ast.Name):
                name = fn.id
            elif isinstance(fn, ast.Attribute):
                name = fn.attr
            if name not in KEY_CALLS:
                continue
            if node.args and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
                keys.add(node.args[0].value)
    # shared validation-batch keys emitted from common helper for module status/activity lines
    keys.update(
        {
            "validation.batch.title_success",
            "validation.batch.title_error",
            "validation.batch.accepted_count",
            "validation.batch.rejected_count",
            "validation.batch.details_prefix",
            "validation.batch.detail_entry",
            "validation.batch.details_entries_1",
            "validation.batch.details_entries_2",
            "validation.batch.details_entries_3",
            "validation.batch.details_entries_3_more",
            "validation.batch.more_suffix",
            "validation.batch.activity_line",
            "validation.batch.activity_segment",
            "common.validation_failed_invalid_data",
            "common.error_while_process",
        }
    )
    return keys


def _load_locale_map(locale_file: str) -> dict[str, str]:
    """Load locale map.
    
    Args:
        locale_file: Parameter value (str).
    
    Returns:
        dict[str, str]: Return value.
    
    Raises:
        None.
    """
    raw = (REPO_ROOT / "common" / "i18n" / locale_file).read_text(encoding="utf-8")
    mapping: dict[str, str] = {}
    for message_xml in re.findall(r"<message\b.*?</message>", raw, flags=re.DOTALL):
        source_match = re.search(r"<source>(.*?)</source>", message_xml, flags=re.DOTALL)
        translation_match = re.search(r"<translation(?:\s+[^>]*)?>(.*?)</translation>", message_xml, flags=re.DOTALL)
        if source_match is None or translation_match is None:
            continue
        source = html.unescape(source_match.group(1).strip())
        translation = html.unescape(re.sub(r"<[^>]+>", "", translation_match.group(1)).strip())
        mapping[source] = translation
    return mapping


def test_module_emitted_i18n_keys_have_non_key_translations_in_all_locales() -> None:
    """Test module emitted i18n keys have non key translations in all locales.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    keys = sorted(_collect_module_i18n_keys())
    violations: list[str] = []
    for locale in LOCALES:
        mapping = _load_locale_map(locale)
        for key in keys:
            translation = mapping.get(key)
            if translation is None:
                violations.append(f"{locale}: missing key '{key}'")
                continue
            if translation.strip() == key.strip():
                violations.append(f"{locale}: key-echo translation '{key}'")
    if not (not violations):
        raise AssertionError("Locale coverage violations:\n" + "\n".join(violations))

