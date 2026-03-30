from __future__ import annotations

import pytest

from common.exceptions import ValidationError
from domain import template_strategy_router as router


def test_router_accepts_marks_template_batch_validation(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, object] = {}

    class _DummyStrategy:
        def validate_workbooks(
            self,
            *,
            template_id: str,
            workbook_kind: str,
            workbook_paths: list[str],
            cancel_token: object | None = None,
            context: object | None = None,
        ) -> dict[str, object]:
            del cancel_token
            del context
            captured["template_id"] = template_id
            captured["workbook_kind"] = workbook_kind
            captured["workbook_paths"] = workbook_paths
            return {"valid_paths": list(workbook_paths)}

    monkeypatch.setattr(router, "get_template_strategy", lambda _template_id: _DummyStrategy())
    result = router.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_paths=["marks.xlsx"],
        workbook_kind="marks_template",
    )

    assert result["valid_paths"] == ["marks.xlsx"]
    assert captured["workbook_kind"] == "marks_template"


def test_router_rejects_unsupported_batch_validation_kind() -> None:
    with pytest.raises(ValidationError) as excinfo:
        router.validate_workbooks(
            template_id="COURSE_SETUP_V2",
            workbook_paths=["marks.xlsx"],
            workbook_kind="unsupported_kind",
        )
    assert getattr(excinfo.value, "code", None) == "WORKBOOK_KIND_UNSUPPORTED"
