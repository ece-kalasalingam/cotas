from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication, QFileDialog

from common.utils import canonical_path_key
from modules import instructor_module as instructor_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _dispose_widget(widget: object, qapp: QApplication) -> None:
    close = getattr(widget, "close", None)
    if callable(close):
        close()
    delete_later = getattr(widget, "deleteLater", None)
    if callable(delete_later):
        delete_later()
    qapp.processEvents()


def _build_instructor_module(monkeypatch: pytest.MonkeyPatch) -> instructor_ui.InstructorModule:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    return instructor_ui.InstructorModule()


def _run_async_inline(
    module: instructor_ui.InstructorModule,
    *,
    token: object,
    job_id: str | None,
    work: Any,
    on_success: Any,
    on_failure: Any,
    on_finally: Any = None,
) -> None:
    del module, token, job_id
    try:
        result = work()
        on_success(result)
    except Exception as exc:  # pragma: no cover - defensive in test shim
        on_failure(exc)
    finally:
        if callable(on_finally):
            on_finally()


def test_marks_generation_retries_only_collisions_with_output_overrides(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication, tmp_path: Path
) -> None:
    module = _build_instructor_module(monkeypatch)
    src1 = "C:/src/one.xlsx"
    src2 = "C:/src/two.xlsx"
    module.course_details_paths = [src1, src2]
    module._validated_template_ids_by_path_key = {
        canonical_path_key(src1): "COURSE_SETUP_V2",
    }

    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *_a, **_k: str(tmp_path)))
    monkeypatch.setattr(
        instructor_ui.InstructorModule,
        "_start_async_operation",
        _run_async_inline,
    )

    prompts: list[str] = []
    monkeypatch.setattr(
        module,
        "_prompt_output_path_for_collision",
        lambda suggested: prompts.append(str(suggested)) or str(tmp_path / "renamed_two.xlsx"),
    )
    monkeypatch.setattr(module, "_prompt_overwrite_all_conflicts", lambda _paths: False)

    calls: list[dict[str, object]] = []

    def _fake_generate_workbooks(**kwargs: object) -> dict[str, object]:
        calls.append(dict(kwargs))
        run_sources = list(kwargs.get("workbook_paths", []))
        context = dict(kwargs.get("context", {}))
        if run_sources == [src1, src2] and context.get("overwrite_existing") is False:
            return {
                "total": 2,
                "generated": 1,
                "failed": 1,
                "skipped": 0,
                "generated_workbook_paths": [str(tmp_path / "one_marks.xlsx")],
                "results": {
                    "s1": {
                        "status": "generated",
                        "source_path": src1,
                        "workbook_path": str(tmp_path / "one_marks.xlsx"),
                        "reason": None,
                    },
                    "s2": {
                        "status": "failed",
                        "source_path": src2,
                        "output_path": str(tmp_path / "two_marks.xlsx"),
                        "existing_output_path": str(tmp_path / "two_marks.xlsx"),
                        "reason": "output_already_exists",
                    },
                },
            }
        assert run_sources == [src2]
        assert context.get("overwrite_existing") is True
        assert context.get("output_path_overrides") == {src2: str(tmp_path / "renamed_two.xlsx")}
        return {
            "total": 1,
            "generated": 1,
            "failed": 0,
            "skipped": 0,
            "generated_workbook_paths": [str(tmp_path / "renamed_two.xlsx")],
            "results": {
                "s2": {
                    "status": "generated",
                    "source_path": src2,
                    "workbook_path": str(tmp_path / "renamed_two.xlsx"),
                    "reason": None,
                }
            },
        }

    monkeypatch.setattr(instructor_ui, "generate_workbooks", _fake_generate_workbooks)

    module._prepare_marks_template_async()

    assert len(calls) == 2
    assert prompts == [str(tmp_path / "two_marks.xlsx")]
    assert module.marks_template_paths == [
        str(tmp_path / "one_marks.xlsx"),
        str(tmp_path / "renamed_two.xlsx"),
    ]
    assert module.marks_template_path == str(tmp_path / "renamed_two.xlsx")
    _dispose_widget(module, qapp)


def test_marks_generation_uses_bulk_overwrite_prompt_when_collisions_exceed_limit(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication, tmp_path: Path
) -> None:
    module = _build_instructor_module(monkeypatch)
    source_paths = [f"C:/src/{idx}.xlsx" for idx in range(1, 8)]
    module.course_details_paths = list(source_paths)
    module._validated_template_ids_by_path_key = {
        canonical_path_key(source_paths[0]): "COURSE_SETUP_V2",
    }

    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *_a, **_k: str(tmp_path)))
    monkeypatch.setattr(
        instructor_ui.InstructorModule,
        "_start_async_operation",
        _run_async_inline,
    )

    bulk_prompts: list[list[str]] = []
    per_file_prompts: list[str] = []
    monkeypatch.setattr(
        module,
        "_prompt_overwrite_all_conflicts",
        lambda paths: bulk_prompts.append(list(paths)) or True,
    )
    monkeypatch.setattr(
        module,
        "_prompt_output_path_for_collision",
        lambda suggested: per_file_prompts.append(str(suggested)) or None,
    )

    call_index = {"value": 0}

    def _fake_generate_workbooks(**kwargs: object) -> dict[str, object]:
        call_index["value"] += 1
        context = dict(kwargs.get("context", {}))
        run_sources = list(kwargs.get("workbook_paths", []))
        if call_index["value"] == 1:
            assert context.get("overwrite_existing") is False
            results = {}
            for source in run_sources:
                out = str(tmp_path / f"{Path(source).stem}_marks.xlsx")
                results[source] = {
                    "status": "failed",
                    "source_path": source,
                    "output_path": out,
                    "existing_output_path": out,
                    "reason": "output_already_exists",
                }
            return {
                "total": len(run_sources),
                "generated": 0,
                "failed": len(run_sources),
                "skipped": 0,
                "generated_workbook_paths": [],
                "results": results,
            }

        assert context.get("overwrite_existing") is True
        overrides = dict(context.get("output_path_overrides", {}))
        assert set(overrides.keys()) == set(source_paths)
        generated_paths = [str(overrides[src]) for src in source_paths]
        return {
            "total": len(run_sources),
            "generated": len(run_sources),
            "failed": 0,
            "skipped": 0,
            "generated_workbook_paths": generated_paths,
            "results": {
                src: {
                    "status": "generated",
                    "source_path": src,
                    "workbook_path": str(overrides[src]),
                    "reason": None,
                }
                for src in source_paths
            },
        }

    monkeypatch.setattr(instructor_ui, "generate_workbooks", _fake_generate_workbooks)

    module._prepare_marks_template_async()

    assert call_index["value"] == 2
    assert len(bulk_prompts) == 1
    assert len(bulk_prompts[0]) == len(source_paths)
    assert per_file_prompts == []
    assert len(module.marks_template_paths) == len(source_paths)
    _dispose_widget(module, qapp)


def test_course_details_validation_uses_shared_batch_feedback(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    module = _build_instructor_module(monkeypatch)
    monkeypatch.setattr(
        instructor_ui.InstructorModule,
        "_start_async_operation",
        _run_async_inline,
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_workbooks",
        lambda **kwargs: {
            "valid_paths": ["C:/valid.xlsx"],
            "invalid_paths": ["C:/invalid.xlsx"],
            "mismatched_paths": [],
            "duplicate_paths": [],
            "duplicate_sections": [],
            "rejections": [
                {
                    "path": "C:/invalid.xlsx",
                    "issue": {
                        "code": "COURSE_DETAILS_COHORT_MISMATCH",
                        "translation_key": "validation.course_details.cohort_mismatch",
                        "message": "mismatch",
                        "context": {"workbook": "C:/invalid.xlsx", "fields": "Course_Code"},
                    },
                }
            ],
            "total": len(kwargs.get("workbook_paths", [])),
        },
    )

    module._upload_course_details_from_paths_async(["C:/valid.xlsx", "C:/invalid.xlsx"])

    assert any(
        "validation.batch.title_error" in str(entry.get("message", ""))
        for entry in getattr(module, "_user_log_entries", [])
    )
    _dispose_widget(module, qapp)
