from __future__ import annotations

from pathlib import Path

import pytest

from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2 import CourseSetupV2Strategy


def test_strategy_generate_workbook_routes_co_attainment_to_generator(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test strategy generate workbook routes co attainment to generator.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()
    captured: dict[str, object] = {}

    class _Signature:
        template_id = "COURSE_SETUP_V2"
        total_outcomes = 6

    def _fake_inputs(*, context, output_path, default_template_id):
        """Fake inputs.
        
        Args:
            context: Parameter value.
            output_path: Parameter value.
            default_template_id: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del context
        return {
            "source_paths": [tmp_path / "source.xlsx"],
            "output_path": output_path,
            "template_id": default_template_id,
            "total_outcomes": _Signature.total_outcomes,
            "thresholds": (40.0, 60.0, 75.0),
            "co_attainment_percent": 80.0,
            "co_attainment_level": 2,
        }

    def _fake_generator():
        """Fake generator.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        def _impl(**kwargs):
            """Impl.
            
            Args:
                kwargs: Parameter value.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            captured.update(kwargs)
            return kwargs["output_path"]

        return _impl

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.co_attainment_generation_inputs",
        _fake_inputs,
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.co_attainment_generator",
        _fake_generator,
    )

    output_path = tmp_path / "co_attainment.xlsx"
    result = strategy.generate_workbook(
        template_id="COURSE_SETUP_V2",
        workbook_kind="co_attainment",
        output_path=output_path,
        workbook_name=output_path.name,
        context={
            "source_paths": [str(tmp_path / "source.xlsx")],
            "thresholds": (40.0, 60.0, 75.0),
            "co_attainment_percent": 80.0,
            "co_attainment_level": 2,
        },
    )

    assert str(result) == str(output_path)
    assert captured["template_id"] == "COURSE_SETUP_V2"
    assert captured["source_paths"] == [tmp_path / "source.xlsx"]
    assert captured["co_attainment_percent"] == 80.0
    assert captured["co_attainment_level"] == 2


def test_strategy_generate_workbook_co_attainment_requires_source_paths(tmp_path: Path) -> None:
    """Test strategy generate workbook co attainment requires source paths.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()
    output_path = tmp_path / "co_attainment.xlsx"

    with pytest.raises(ValidationError) as excinfo:
        strategy.generate_workbook(
            template_id="COURSE_SETUP_V2",
            workbook_kind="co_attainment",
            output_path=output_path,
            workbook_name=output_path.name,
            context={},
        )

    assert getattr(excinfo.value, "code", None) == "COA_SOURCE_WORKBOOK_REQUIRED"


def test_strategy_extract_course_metadata_and_students_routes_to_v2_extractor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test strategy extract course metadata and students routes to v2 extractor.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()
    captured: dict[str, object] = {}

    def _fake_extractor():
        """Fake extractor.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        def _impl(path: Path) -> tuple[set[str], dict[str, str]]:
            """Impl.
            
            Args:
                path: Parameter value (Path).
            
            Returns:
                tuple[set[str], dict[str, str]]: Return value.
            
            Raises:
                None.
            """
            captured["path"] = path
            return {"r1"}, {"course_code": "CSE101"}

        return _impl

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.course_metadata_students_extractor",
        _fake_extractor,
    )

    workbook_path = tmp_path / "marks.xlsx"
    students, metadata = strategy.extract_course_metadata_and_students(
        workbook_path,
        template_id="COURSE_SETUP_V2",
    )

    assert captured["path"] == workbook_path
    assert students == {"r1"}
    assert metadata == {"course_code": "CSE101"}
