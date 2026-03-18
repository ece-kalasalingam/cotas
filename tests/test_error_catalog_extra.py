from common import error_catalog


def test_resolve_validation_error_message_returns_key_on_translation_error(monkeypatch) -> None:
    monkeypatch.setattr(error_catalog, "t", lambda *_a, **_k: (_ for _ in ()).throw(RuntimeError("boom")))
    assert (
        error_catalog.resolve_validation_error_message("WORKBOOK_NOT_FOUND", {"workbook": "x.xlsx"})
        == "instructor.validation.workbook_not_found"
    )
