import logging

import common.texts as texts
from common.texts import get_available_languages, set_language, set_language_from_system, t
from common.texts.en import TEXTS as EN_TEXTS
from common.texts.hi_in import TEXTS as HI_IN_TEXTS
from common.texts.ta_in import TEXTS as TA_IN_TEXTS
from common.texts.te_in import TEXTS as TE_IN_TEXTS


def test_t_formats_placeholders():
    assert t("about.version", version="1.2.3.0") == "Version 1.2.3.0"


def test_all_language_catalogs_have_matching_key_counts_and_sets():
    catalogs = {
        "en": EN_TEXTS,
        "ta-in": TA_IN_TEXTS,
        "hi-in": HI_IN_TEXTS,
        "te-in": TE_IN_TEXTS,
    }
    expected_keys = set(EN_TEXTS)
    expected_count = len(EN_TEXTS)
    for code, catalog in catalogs.items():
        assert len(catalog) == expected_count, f"{code} key count mismatch"
        assert set(catalog) == expected_keys, f"{code} key set mismatch"


def test_unknown_language_falls_back_to_default():
    set_language("unknown")
    assert t("status.ready") == "Ready"
    set_language("en")


def test_tamil_india_language_works():
    set_language("ta-IN")
    assert t("status.ready") == "தயார்"
    set_language("en")


def test_tamil_india_underscore_variant_works():
    set_language("ta_IN")
    assert t("status.ready") == "தயார்"
    set_language("en")


def test_hindi_india_language_works():
    set_language("hi-IN")
    assert t("status.ready") == "तैयार"
    set_language("en")


def test_hindi_india_underscore_variant_works():
    set_language("hi_IN")
    assert t("status.ready") == "तैयार"
    set_language("en")


def test_telugu_india_language_works():
    set_language("te-IN")
    assert t("status.ready") == "సిద్ధం"
    set_language("en")


def test_telugu_india_underscore_variant_works():
    set_language("te_IN")
    assert t("status.ready") == "సిద్ధం"
    set_language("en")


def test_system_language_uses_windows_lcid_when_available():
    set_language_from_system(system_lcid=1097, system_locale="en_US")
    assert t("status.ready") == "தயார்"
    set_language("en")


def test_system_language_uses_windows_lcid_for_hindi_when_available():
    set_language_from_system(system_lcid=1081, system_locale="en_US")
    assert t("status.ready") == "तैयार"
    set_language("en")


def test_system_language_uses_windows_lcid_for_telugu_when_available():
    set_language_from_system(system_lcid=1098, system_locale="en_US")
    assert t("status.ready") == "సిద్ధం"
    set_language("en")


def test_system_language_falls_back_to_english_1033():
    set_language_from_system(system_lcid=9999, system_locale="xx_YY")
    assert t("status.ready") == "Ready"
    set_language("en")


def test_available_language_labels_are_native_and_language_independent():
    set_language("en")
    labels_in_english_ui = dict(get_available_languages())

    set_language("ta-IN")
    labels_in_tamil_ui = dict(get_available_languages())

    assert labels_in_english_ui == labels_in_tamil_ui
    assert labels_in_english_ui["en"] == "English"
    assert labels_in_english_ui["hi-in"] == "हिन्दी (भारत)"
    assert labels_in_english_ui["ta-in"] == "தமிழ் (இந்தியா)"
    assert labels_in_english_ui["te-in"] == "తెలుగు (భారతదేశం)"
    set_language("en")


def test_mojibake_catalog_falls_back_to_english_with_warning(caplog):
    original_catalog = texts._CATALOGS["ta-in"]
    original_cache = dict(texts._catalog_health_cache)
    original_warned = set(texts._warned_bad_catalogs)

    try:
        texts._CATALOGS["ta-in"] = {
            "status.ready": "à®¤à®¯à®¾à®°à¯",
            "about.version": "à®ªà®¤à®¿à®ªà¯à®ªà¯ {version}",
        }
        texts._catalog_health_cache.clear()
        texts._warned_bad_catalogs.clear()

        with caplog.at_level(logging.WARNING):
            texts.set_language("ta-IN")

        assert texts.t("status.ready") == "Ready"
        assert "Detected mojibake-like UI strings" in caplog.text
    finally:
        texts._CATALOGS["ta-in"] = original_catalog
        texts._catalog_health_cache.clear()
        texts._catalog_health_cache.update(original_cache)
        texts._warned_bad_catalogs.clear()
        texts._warned_bad_catalogs.update(original_warned)
        texts.set_language("en")


def test_set_language_accepts_short_aliases() -> None:
    set_language("hi")
    assert t("status.ready") == "तैयार"

    set_language("ta")
    assert t("status.ready") == "தயார்"

    set_language("te")
    assert t("status.ready") == "సిద్ధం"
    set_language("en")


def test_get_windows_ui_lcid_non_windows_and_exception(monkeypatch) -> None:
    monkeypatch.setattr(texts.sys, "platform", "linux")
    assert texts._get_windows_ui_lcid() is None

    monkeypatch.setattr(texts.sys, "platform", "win32")

    class _BrokenFn:
        def __call__(self):
            raise RuntimeError("boom")

    monkeypatch.setattr(
        texts.ctypes,
        "windll",
        type("_Windll", (), {"kernel32": type("_K32", (), {"GetUserDefaultUILanguage": _BrokenFn()})()})(),
        raising=False,
    )
    assert texts._get_windows_ui_lcid() is None


def test_catalog_health_false_for_missing_catalog() -> None:
    texts._catalog_health_cache.clear()
    assert texts._catalog_is_healthy("xx") is False
    assert texts._catalog_health_cache.get("xx") is False


def test_system_language_uses_locale_getlocale_when_inputs_missing(monkeypatch) -> None:
    monkeypatch.setattr(texts, "_get_windows_ui_lcid", lambda: None)
    monkeypatch.setattr(texts.locale, "getlocale", lambda: ("te_IN", "UTF-8"))

    set_language_from_system(system_lcid=None, system_locale=None)
    assert t("status.ready") == "సిద్ధం"
    set_language("en")


def test_system_language_handles_language_only_locales() -> None:
    set_language_from_system(system_lcid=None, system_locale="hi")
    assert t("status.ready") == "तैयार"

    set_language_from_system(system_lcid=None, system_locale="ta")
    assert t("status.ready") == "தயார்"

    set_language_from_system(system_lcid=None, system_locale="te")
    assert t("status.ready") == "సిద్ధం"

    set_language_from_system(system_lcid=None, system_locale="en")
    assert t("status.ready") == "Ready"


def test_t_falls_back_to_key_when_missing_in_all_catalogs() -> None:
    set_language("en")
    assert t("__definitely_missing_key__") == "__definitely_missing_key__"


def test_system_language_maps_regional_variants_via_lang_only_fallbacks() -> None:
    set_language_from_system(system_lcid=None, system_locale="hi-XX")
    assert t("status.ready") == "तैयार"

    set_language_from_system(system_lcid=None, system_locale="ta-XX")
    assert t("status.ready") == "தயார்"

    set_language_from_system(system_lcid=None, system_locale="te-XX")
    assert t("status.ready") == "సిద్ధం"

    set_language_from_system(system_lcid=None, system_locale="en-GB")
    assert t("status.ready") == "Ready"
    set_language("en")
