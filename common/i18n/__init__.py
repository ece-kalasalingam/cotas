"""Qt translator accessors for user-facing UI strings."""

from __future__ import annotations

import locale
import os
import sys
from pathlib import Path
from typing import Any

from PySide6.QtCore import QCoreApplication, QLocale, QTranslator

from common.exceptions import ConfigurationError

_TRANSLATION_CONTEXT = "main"
_DEFAULT_LANG = "en"
_I18N_RELATIVE_DIR = "common/i18n"
_QM_BASE_NAME = "obe"

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

_LANG_TO_QT_LOCALE: dict[str, str] = {
    "en": "en_US",
    "hi-in": "hi_IN",
    "ta-in": "ta_IN",
    "te-in": "te_IN",
}

_active_lang = _DEFAULT_LANG
_active_translator: QTranslator | None = None


def _normalize_lang(lang: str) -> str:
    code = (lang or "").strip().replace("_", "-").lower()
    if code == "hi":
        return "hi-in"
    if code == "ta":
        return "ta-in"
    if code == "te":
        return "te-in"
    return code


def _resolve_supported_language(lang: str) -> str:
    normalized = _normalize_lang(lang)
    if normalized in _LANG_TO_QT_LOCALE:
        return normalized
    return _DEFAULT_LANG


def _get_windows_ui_lcid() -> int | None:
    if not sys.platform.startswith("win"):
        return None
    try:
        import ctypes
        from ctypes import wintypes

        windll = getattr(ctypes, "windll", None)
        if windll is None:
            return None
        kernel32: Any = windll.kernel32
        get_ui_lang = kernel32.GetUserDefaultUILanguage
        get_ui_lang.argtypes = []
        get_ui_lang.restype = wintypes.USHORT
        return int(get_ui_lang())
    except Exception:
        return None


def _i18n_dir() -> str:
    base = getattr(
        sys,
        "_MEIPASS",
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    )
    return str(Path(os.path.join(base, _I18N_RELATIVE_DIR)))


def _install_translator_for_language(lang: str) -> None:
    global _active_translator
    app = QCoreApplication.instance()
    if app is None:
        return

    if _active_translator is not None:
        app.removeTranslator(_active_translator)
        _active_translator = None

    qt_locale_name = _LANG_TO_QT_LOCALE[lang]
    translator = QTranslator()
    if not translator.load(QLocale(qt_locale_name), _QM_BASE_NAME, "_", _i18n_dir()):
        raise ConfigurationError(
            f"Qt translation catalog load failed for language '{lang}' (locale '{qt_locale_name}')."
        )

    app.installTranslator(translator)
    _active_translator = translator


def set_language(lang: str) -> None:
    """Set active language and install Qt translator if app exists."""
    global _active_lang
    _active_lang = _resolve_supported_language(lang)
    _install_translator_for_language(_active_lang)


def set_language_from_system(
    *, system_lcid: int | None = None, system_locale: str | None = None
) -> None:
    """Set language from OS/user locale, fallback to English."""
    if system_lcid is not None:
        mapped = _LCID_TO_LANG.get(system_lcid)
        if mapped in _LANG_TO_QT_LOCALE:
            set_language(mapped)
            return

    locale_code = system_locale
    if not locale_code:
        locale_code = locale.getlocale()[0]
    if locale_code:
        normalized = _normalize_lang(locale_code)
        if normalized in _LANG_TO_QT_LOCALE:
            set_language(normalized)
            return
        lang_only = normalized.split("-")[0]
        if lang_only == "hi":
            set_language("hi-in")
            return
        if lang_only == "ta":
            set_language("ta-in")
            return
        if lang_only == "te":
            set_language("te-in")
            return
        if lang_only == "en":
            set_language("en")
            return

    if system_lcid is None and system_locale is None:
        lcid = _get_windows_ui_lcid()
        if lcid is not None:
            mapped = _LCID_TO_LANG.get(lcid)
            if mapped in _LANG_TO_QT_LOCALE:
                set_language(mapped)
                return

    set_language(_DEFAULT_LANG)


def get_language() -> str:
    """Return active language code."""
    return _active_lang


def get_available_languages() -> tuple[tuple[str, str], ...]:
    """Return supported language options as (code, label)."""
    return tuple((code, _LANGUAGE_LABELS.get(code, code)) for code in _LANG_TO_QT_LOCALE)


def t(key: str, **kwargs: object) -> str:
    """Translate a UI text key through Qt's translation system."""
    translated = QCoreApplication.translate(_TRANSLATION_CONTEXT, key)
    return translated.format(**kwargs) if kwargs else translated
