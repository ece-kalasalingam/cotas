from __future__ import annotations

from pathlib import Path

import pytest

from common.exceptions import ValidationError
from services import co_analysis_workflow_service as service_mod


def test_generate_workbook_validates_all_sources_before_generation(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    service = service_mod.CoAnalysisWorkflowService()
    context = service.create_job_context(step_id="generate")
    source_paths = [tmp_path / "a.xlsx", tmp_path / "b.xlsx"]
    output = tmp_path / "out.xlsx"

    seen: dict[str, object] = {"validated": [], "generated": False}

    def _fake_validate(path: Path) -> None:
        cast_list = seen["validated"]
        assert isinstance(cast_list, list)
        cast_list.append(path)

    def _fake_generate(*, source_paths, output_path, token, **kwargs):
        del token
        del kwargs
        seen["generated"] = True
        return output_path

    monkeypatch.setattr(service_mod, "validate_uploaded_source_workbook", _fake_validate)
    monkeypatch.setattr(service_mod, "consume_last_source_anomaly_warnings", lambda: [])
    monkeypatch.setattr(service_mod, "generate_co_analysis_workbook", _fake_generate)

    result = service.generate_workbook(
        source_paths=source_paths,
        output_path=output,
        context=context,
    )

    assert result == output
    assert seen["validated"] == source_paths
    assert seen["generated"] is True


def test_generate_workbook_stops_when_validation_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    service = service_mod.CoAnalysisWorkflowService()
    context = service.create_job_context(step_id="generate")
    source_paths = [tmp_path / "a.xlsx", tmp_path / "b.xlsx"]
    output = tmp_path / "out.xlsx"

    seen = {"generated": False}

    def _fake_validate(path: Path) -> None:
        if path.name == "b.xlsx":
            raise ValidationError("invalid", code="TEST_INVALID")

    def _fake_generate(*, source_paths, output_path, token, **kwargs):
        del source_paths
        del output_path
        del token
        del kwargs
        seen["generated"] = True
        return output

    monkeypatch.setattr(service_mod, "validate_uploaded_source_workbook", _fake_validate)
    monkeypatch.setattr(service_mod, "consume_last_source_anomaly_warnings", lambda: [])
    monkeypatch.setattr(service_mod, "generate_co_analysis_workbook", _fake_generate)

    with pytest.raises(ValidationError, match="invalid"):
        service.generate_workbook(
            source_paths=source_paths,
            output_path=output,
            context=context,
        )

    assert seen["generated"] is False
