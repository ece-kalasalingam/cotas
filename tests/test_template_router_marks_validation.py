from __future__ import annotations

import pytest

from common.exceptions import ValidationError
from domain import template_strategy_router as router


def test_router_accepts_marks_template_batch_validation(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test router accepts marks template batch validation.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
            """Validate workbooks.
            
            Args:
                template_id: Parameter value (str).
                workbook_kind: Parameter value (str).
                workbook_paths: Parameter value (list[str]).
                cancel_token: Parameter value (object | None).
                context: Parameter value (object | None).
            
            Returns:
                dict[str, object]: Return value.
            
            Raises:
                None.
            """
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


def test_router_accepts_co_description_batch_validation(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test router accepts CO-description template batch validation."""
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
        workbook_paths=["co_description.xlsx"],
        workbook_kind="co_description",
    )

    assert result["valid_paths"] == ["co_description.xlsx"]
    assert captured["workbook_kind"] == "co_description"


def test_router_rejects_unsupported_batch_validation_kind() -> None:
    """Test router rejects unsupported batch validation kind.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    with pytest.raises(ValidationError) as excinfo:
        router.validate_workbooks(
            template_id="COURSE_SETUP_V2",
            workbook_paths=["marks.xlsx"],
            workbook_kind="unsupported_kind",
        )
    assert getattr(excinfo.value, "code", None) == "WORKBOOK_KIND_UNSUPPORTED"


def test_router_consume_marks_anomaly_warnings_uses_strategy(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test router consume marks anomaly warnings uses strategy.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    class _DummyStrategy:
        def consume_last_marks_anomaly_warnings(self) -> list[str]:
            """Consume last marks anomaly warnings.
            
            Args:
                None.
            
            Returns:
                list[str]: Return value.
            
            Raises:
                None.
            """
            return ["warn-a", "", "warn-b"]

    monkeypatch.setattr(router, "get_template_strategy", lambda _template_id: _DummyStrategy())
    warnings = router.consume_marks_anomaly_warnings("COURSE_SETUP_V2")
    assert warnings == ["warn-a", "warn-b"]
