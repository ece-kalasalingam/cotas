"""Simple text catalog accessors for UI strings."""

from __future__ import annotations

from collections.abc import Mapping
import ctypes
import locale
import logging
import sys

from common.texts.en import TEXTS as EN_TEXTS
from common.texts.hi_in import TEXTS as HI_IN_TEXTS
from common.texts.ta_in import TEXTS as TA_IN_TEXTS
from common.texts.te_in import TEXTS as TE_IN_TEXTS

_CATALOGS: dict[str, Mapping[str, str]] = {
    "en": EN_TEXTS,
    "hi-in": HI_IN_TEXTS,
    "ta-in": TA_IN_TEXTS,
    "te-in": TE_IN_TEXTS,
}
_DEFAULT_LANG = "en"
_active_lang = _DEFAULT_LANG
_logger = logging.getLogger(__name__)
_warned_bad_catalogs: set[str] = set()
_catalog_health_cache: dict[str, bool] = {}
_LCID_TO_LANG: dict[int, str] = {
    1033: "en",  # English (United States)
    1081: "hi-in",  # Hindi (India)
    1097: "ta-in",  # Tamil (India)
    1098: "te-in",  # Telugu (India)
}
_LANGUAGE_LABELS: dict[str, str] = {
    "en": "English",
    "hi-in": "हिन्दी (भारत)",
    "ta-in": "தமிழ் (இந்தியா)",
    "te-in": "తెలుగు (భారతదేశం)",
}


def _normalize_lang(lang: str) -> str:
    code = (lang or "").strip().replace("_", "-").lower()
    if code == "hi":
        return "hi-in"
    if code == "ta":
        return "ta-in"
    if code == "te":
        return "te-in"
    return code


def _get_windows_ui_lcid() -> int | None:
    if not sys.platform.startswith("win"):
        return None
    try:
        from ctypes import wintypes

        get_ui_lang = ctypes.windll.kernel32.GetUserDefaultUILanguage
        get_ui_lang.argtypes = []
        get_ui_lang.restype = wintypes.USHORT
        return int(get_ui_lang())
    except Exception:
        return None


def _looks_like_mojibake(value: str) -> bool:
    # Common UTF-8->Latin1 mojibake markers seen in UI catalogs.
    markers = ("à®", "à¯", "Ã", "Â", "Ð", "Ñ")
    return any(marker in value for marker in markers)


def _catalog_is_healthy(lang: str) -> bool:
    cached = _catalog_health_cache.get(lang)
    if cached is not None:
        return cached

    catalog = _CATALOGS.get(lang)
    if not catalog:
        _catalog_health_cache[lang] = False
        return False

    for value in catalog.values():
        if isinstance(value, str) and _looks_like_mojibake(value):
            if lang not in _warned_bad_catalogs:
                _logger.warning(
                    "Detected mojibake-like UI strings for '%s'; falling back to '%s'.",
                    lang,
                    _DEFAULT_LANG,
                )
                _warned_bad_catalogs.add(lang)
            _catalog_health_cache[lang] = False
            return False

    _catalog_health_cache[lang] = True
    return True


def set_language(lang: str) -> None:
    """Set active language catalog."""
    global _active_lang
    normalized = _normalize_lang(lang)
    if normalized in _CATALOGS and _catalog_is_healthy(normalized):
        _active_lang = normalized
        return
    _active_lang = _DEFAULT_LANG


def set_language_from_system(
    *, system_lcid: int | None = None, system_locale: str | None = None
) -> None:
    """Set language from OS/user locale, fallback to English (LCID 1033)."""
    global _active_lang

    lcid = system_lcid if system_lcid is not None else _get_windows_ui_lcid()
    if lcid is not None:
        mapped = _LCID_TO_LANG.get(lcid)
        if mapped in _CATALOGS and _catalog_is_healthy(mapped):
            _active_lang = mapped
            return

    locale_code = system_locale
    if not locale_code:
        locale_code = locale.getlocale()[0]
    if locale_code:
        normalized = _normalize_lang(locale_code)
        if normalized in _CATALOGS and _catalog_is_healthy(normalized):
            _active_lang = normalized
            return
        # Handle language-only locales such as "ta".
        lang_only = normalized.split("-")[0]
        if (
            lang_only == "hi"
            and "hi-in" in _CATALOGS
            and _catalog_is_healthy("hi-in")
        ):
            _active_lang = "hi-in"
            return
        if (
            lang_only == "ta"
            and "ta-in" in _CATALOGS
            and _catalog_is_healthy("ta-in")
        ):
            _active_lang = "ta-in"
            return
        if (
            lang_only == "te"
            and "te-in" in _CATALOGS
            and _catalog_is_healthy("te-in")
        ):
            _active_lang = "te-in"
            return
        if lang_only == "en" and "en" in _CATALOGS and _catalog_is_healthy("en"):
            _active_lang = "en"
            return

    # Default to English (Windows 1033 equivalent).
    _active_lang = "en"


def get_language() -> str:
    """Return active language code."""
    return _active_lang


def get_available_languages() -> tuple[tuple[str, str], ...]:
    """Return supported language options as (code, label)."""
    return tuple((code, _LANGUAGE_LABELS.get(code, code)) for code in _CATALOGS)


def t(key: str, **kwargs: object) -> str:
    """Lookup a text key and format placeholders."""
    template = _CATALOGS.get(_active_lang, _CATALOGS[_DEFAULT_LANG]).get(key)
    if template is None:
        template = _CATALOGS[_DEFAULT_LANG].get(key, key)
    return template.format(**kwargs) if kwargs else template
