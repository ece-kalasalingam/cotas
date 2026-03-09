import logging

import common.texts as texts
from common.texts import get_available_languages, set_language, set_language_from_system, t


def test_t_formats_placeholders():
    assert t("about.version", version="1.2.3.0") == "Version 1.2.3.0"


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


def test_system_language_uses_windows_lcid_when_available():
    set_language_from_system(system_lcid=1097, system_locale="en_US")
    assert t("status.ready") == "தயார்"
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
    assert labels_in_english_ui["ta-in"] == "தமிழ் (இந்தியா)"
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
