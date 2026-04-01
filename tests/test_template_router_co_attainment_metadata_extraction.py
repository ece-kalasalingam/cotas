from __future__ import annotations

from pathlib import Path

from domain import template_strategy_router as router


def test_router_extract_course_metadata_and_students_uses_strategy(monkeypatch) -> None:
    """Test router extract course metadata and students uses strategy.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    captured: dict[str, object] = {}

    class _DummyStrategy:
        def extract_course_metadata_and_students(
            self,
            workbook_path: str | Path,
            *,
            template_id: str,
        ) -> tuple[set[str], dict[str, str]]:
            """Extract course metadata and students.
            
            Args:
                workbook_path: Parameter value (str | Path).
                template_id: Parameter value (str).
            
            Returns:
                tuple[set[str], dict[str, str]]: Return value.
            
            Raises:
                None.
            """
            captured["workbook_path"] = workbook_path
            captured["template_id"] = template_id
            return {"r1", "r2"}, {"course_code": "CSE101"}

    monkeypatch.setattr(
        router,
        "resolve_template_id_from_workbook_path",
        lambda _path: "COURSE_SETUP_V2",
    )
    monkeypatch.setattr(router, "get_template_strategy", lambda _template_id: _DummyStrategy())

    students, metadata = router.extract_course_metadata_and_students_from_workbook_path("x.xlsx")

    assert captured["workbook_path"] == "x.xlsx"
    assert captured["template_id"] == "COURSE_SETUP_V2"
    assert students == {"r1", "r2"}
    assert metadata == {"course_code": "CSE101"}
