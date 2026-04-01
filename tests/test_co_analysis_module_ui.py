from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import co_analysis_module as co_analysis_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    """Qapp.
    
    Args:
        None.
    
    Returns:
        QApplication: Return value.
    
    Raises:
        None.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _dispose_widget(widget: object, qapp: QApplication) -> None:
    """Dispose widget.
    
    Args:
        widget: Parameter value (object).
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    close = getattr(widget, "close", None)
    if callable(close):
        close()
    delete_later = getattr(widget, "deleteLater", None)
    if callable(delete_later):
        delete_later()
    qapp.processEvents()


def _build_module_with_message_capture(
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[co_analysis_ui.COAnalysisModule, list[str]]:
    """Build module with message capture.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        tuple[co_analysis_ui.COAnalysisModule, list[str]]: Return value.
    
    Raises:
        None.
    """
    seen_keys: list[str] = []
    monkeypatch.setattr(co_analysis_ui, "t", lambda key, **kwargs: key)

    def _capture_notify_message_key(self, text_key: str, **kwargs: object) -> None:
        """Capture notify message key.
        
        Args:
            text_key: Parameter value (str).
            kwargs: Parameter value (object).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del self, kwargs
        seen_keys.append(text_key)

    monkeypatch.setattr(
        co_analysis_ui.ModuleRuntime,
        "notify_message_key",
        _capture_notify_message_key,
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "validate_workbooks",
        lambda **kwargs: {
            "valid_paths": list(kwargs.get("workbook_paths", [])),
            "rejections": [],
            "invalid_paths": [],
        },
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "consume_marks_anomaly_warnings",
        lambda _template_id: [],
    )
    monkeypatch.setattr(
        co_analysis_ui.COAnalysisModule, "_setup_ui_logging", lambda self: None
    )
    module = co_analysis_ui.COAnalysisModule()

    def _run_inline(*, token, job_id, work, on_success, on_failure, on_finally=None):
        """Run inline.
        
        Args:
            token: Parameter value.
            job_id: Parameter value.
            work: Parameter value.
            on_success: Parameter value.
            on_failure: Parameter value.
            on_finally: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del token, job_id
        try:
            result = work()
            on_success(result)
        except Exception as exc:  # pragma: no cover - defensive
            on_failure(exc)
        if on_finally is not None:
            on_finally()

    monkeypatch.setattr(module, "_start_async_operation", _run_inline)
    return module, seen_keys


def test_drop_widget_syncs_files_and_clear(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    """Test drop widget syncs files and clear.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx", "C:/b.xlsx"], emit_drop=True)
    assert module._files == [Path("C:/a.xlsx"), Path("C:/b.xlsx")]
    assert module.clear_button.isEnabled() is True
    assert module.calculate_button.isEnabled() is True

    module.drop_widget.clear_button.click()
    qapp.processEvents()
    assert module._files == []
    assert module.clear_button.isEnabled() is False
    assert module.calculate_button.isEnabled() is False
    _dispose_widget(module, qapp)


def test_submit_triggers_async_save_workbook(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    """Test submit triggers async save workbook.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    called = {"count": 0}

    def _fake_prepare_co_analysis_async() -> None:
        """Fake prepare co analysis async.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        called["count"] += 1

    monkeypatch.setattr(module, "_prepare_co_analysis_async", _fake_prepare_co_analysis_async)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module._on_submit_requested()
    assert called["count"] == 1
    _dispose_widget(module, qapp)


def test_submit_shows_threshold_violation_when_invalid(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
) -> None:
    """Test submit shows threshold violation when invalid.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module.threshold_l1_input.setValue(70.0)
    module.threshold_l2_input.setValue(60.0)
    module.threshold_l3_input.setValue(90.0)
    module._on_submit_requested()
    assert "co_analysis.thresholds.invalid_rule" in seen_keys
    _dispose_widget(module, qapp)


def test_download_co_description_template_link_triggers_generation(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test download co description template link triggers generation.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    output_path = tmp_path / "CO_Description_Template.xlsx"
    called: dict[str, object] = {}

    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: (str(output_path), "Excel Files (*.xlsx)"),
    )

    def _fake_generate_workbook(**kwargs):
        """Fake generate workbook.
        
        Args:
            kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        called["workbook_kind"] = kwargs.get("workbook_kind")
        called["template_id"] = kwargs.get("template_id")
        return type(
            "_Result",
            (),
            {
                "output_path": str(output_path),
                "workbook_path": str(output_path),
            },
        )()

    monkeypatch.setattr(co_analysis_ui, "generate_workbook", _fake_generate_workbook)

    def _run_inline(*, token, job_id, work, on_success, on_failure, on_finally=None):
        """Run inline.
        
        Args:
            token: Parameter value.
            job_id: Parameter value.
            work: Parameter value.
            on_success: Parameter value.
            on_failure: Parameter value.
            on_finally: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del token, job_id
        try:
            result = work()
            on_success(result)
        except Exception as exc:  # pragma: no cover - defensive
            on_failure(exc)
        if on_finally is not None:
            on_finally()

    monkeypatch.setattr(module, "_start_async_operation", _run_inline)

    module._download_co_description_template_async()
    qapp.processEvents()

    assert called["workbook_kind"] == "co_description_template"
    assert called["template_id"] == co_analysis_ui.ID_COURSE_SETUP
    assert output_path in module._downloaded_outputs
    assert "co_analysis.status.co_description_template_generated" in seen_keys
    assert "co_analysis.toast.co_description_template_generated" in seen_keys
    _dispose_widget(module, qapp)


def test_marks_upload_validation_emits_issue_and_summary_toast(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
) -> None:
    """Test marks upload validation emits issue and summary toast.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module, seen_keys = _build_module_with_message_capture(monkeypatch)

    def _fake_batch_validate(paths: list[str]) -> dict[str, object]:
        """Fake batch validate.
        
        Args:
            paths: Parameter value (list[str]).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        accepted = [path for path in paths if not Path(path).name.lower().startswith("bad")]
        rejected = [
            {
                "path": path,
                "issue": {
                    "code": "MARKS_TEMPLATE_COHORT_MISMATCH",
                    "category": "validation",
                    "severity": "error",
                    "translation_key": "validation.issue",
                    "message": "mismatch",
                    "context": {"workbook": path, "fields": "Course_Code"},
                },
            }
            for path in paths
            if Path(path).name.lower().startswith("bad")
        ]
        return {"valid_paths": accepted, "rejections": rejected, "invalid_paths": [item["path"] for item in rejected]}

    monkeypatch.setattr(
        co_analysis_ui,
        "validate_workbooks",
        lambda **kwargs: _fake_batch_validate(list(kwargs.get("workbook_paths", []))),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "consume_marks_anomaly_warnings",
        lambda _template_id: [],
    )

    module.drop_widget.add_files(["C:/good.xlsx", "C:/bad.xlsx"], emit_drop=True)
    qapp.processEvents()

    assert any(
        "validation.batch.title_error" in str(entry.get("message", ""))
        for entry in getattr(module, "_user_log_entries", [])
    )
    assert seen_keys == []
    _dispose_widget(module, qapp)

