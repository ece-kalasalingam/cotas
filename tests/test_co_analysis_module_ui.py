from __future__ import annotations

from pathlib import Path
from typing import Callable, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from common.module_messages import NotificationChannel
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
    monkeypatch.setattr(
        co_analysis_ui,
        "resolve_template_id_from_workbook_path",
        lambda _workbook_path: co_analysis_ui.ID_COURSE_SETUP,
    )

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
            "valid_paths": [
                str(path)
                for path in list(kwargs.get("workbook_paths", []))
                if (
                    (
                        str(kwargs.get("workbook_kind", "")).strip() == "marks_template"
                        and "co_description" not in Path(str(path)).name.lower()
                    )
                    or (
                        str(kwargs.get("workbook_kind", "")).strip() == "co_description"
                        and "co_description" in Path(str(path)).name.lower()
                    )
                )
            ],
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
    module.generate_word_report_checkbox.setChecked(False)
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
    module.generate_word_report_checkbox.setChecked(False)
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


def test_word_report_toggle_defaults_enabled(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    """Test word report toggle defaults enabled.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).

    Returns:
        None.

    Raises:
        None.
    """
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    assert module.generate_word_report_checkbox.isChecked() is True
    _dispose_widget(module, qapp)


def test_prepare_co_analysis_passes_word_report_context_and_records_docx_output(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test prepare co analysis passes word report context and records docx output.

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
    module.drop_widget.add_files(
        [
            str(tmp_path / "source.xlsx"),
            str(tmp_path / "co_description.xlsx"),
        ],
        emit_drop=True,
    )
    captured: dict[str, object] = {}
    notify_events: list[tuple[str, dict[str, object]]] = []
    original_notify = cast(Callable[..., None], module._runtime.notify_message_key)

    def _capture_notify(text_key: str, **kwargs: object) -> None:
        notify_events.append((text_key, dict(kwargs)))
        original_notify(text_key, **kwargs)

    monkeypatch.setattr(module._runtime, "notify_message_key", _capture_notify)
    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getExistingDirectory",
        lambda *_args, **_kwargs: str(tmp_path),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "extract_course_metadata_and_students_from_workbook_path",
        lambda _path: (
            set(),
            {
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_COURSE_CODE_KEY): "CSE101",
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_ACADEMIC_YEAR_KEY): "2025-26",
            },
        ),
    )

    def _fake_generate_workbook(**kwargs):
        captured.update(kwargs)
        return type(
            "_Result",
            (),
            {
                "output_path": str(tmp_path / "CSE101_2025-26_CO_Analysis.xlsx"),
                "word_report_path": str(tmp_path / "CSE101_2025-26_CO_Analysis_Report.docx"),
            },
        )()

    monkeypatch.setattr(co_analysis_ui, "generate_workbook", _fake_generate_workbook)
    module._prepare_co_analysis_async()
    qapp.processEvents()

    context = cast(dict[str, object], captured.get("context", {}))
    assert context.get("generate_word_report") is True
    assert str(context.get("co_description_path", "")).endswith("co_description.xlsx")
    assert str(context.get("word_output_path", "")).endswith("CSE101_2025-26_CO_Analysis_Report.docx")
    assert "co_analysis.status.output_generated_excel" in seen_keys
    assert "co_analysis.status.output_generated_word" in seen_keys
    assert "co_analysis.status.word_report_generated" in seen_keys
    output_items = module.get_shared_outputs_data().items
    assert any(item.label_key == "co_analysis.links.generated_excel_output" for item in output_items)
    assert any(item.label_key == "co_analysis.links.generated_word_output" for item in output_items)
    assert any(path.suffix == ".docx" for path in module._downloaded_outputs)
    _dispose_widget(module, qapp)


def test_prepare_co_analysis_word_report_failure_emits_warning_and_keeps_excel_success(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test prepare co analysis word report failure emits warning and keeps excel success.

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
    module.drop_widget.add_files(
        [
            str(tmp_path / "source.xlsx"),
            str(tmp_path / "co_description.xlsx"),
        ],
        emit_drop=True,
    )
    notify_events: list[tuple[str, dict[str, object]]] = []
    original_notify = cast(Callable[..., None], module._runtime.notify_message_key)

    def _capture_notify(text_key: str, **kwargs: object) -> None:
        notify_events.append((text_key, dict(kwargs)))
        original_notify(text_key, **kwargs)

    monkeypatch.setattr(module._runtime, "notify_message_key", _capture_notify)
    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getExistingDirectory",
        lambda *_args, **_kwargs: str(tmp_path),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "extract_course_metadata_and_students_from_workbook_path",
        lambda _path: (set(), {}),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "generate_workbook",
        lambda **_kwargs: type(
            "_Result",
            (),
            {
                "output_path": str(tmp_path / "CO_Analysis.xlsx"),
                "word_report_error_key": "co_analysis.status.word_report_generate_failed",
            },
        )(),
    )
    module._prepare_co_analysis_async()
    qapp.processEvents()

    assert "co_analysis.status.output_generated_excel" in seen_keys
    assert "co_analysis.status.calculate_completed" in seen_keys
    assert "co_analysis.status.word_report_generate_failed" in seen_keys
    output_items = module.get_shared_outputs_data().items
    assert any(item.label_key == "co_analysis.links.generated_excel_output" for item in output_items)
    assert any(path.suffix == ".xlsx" for path in module._downloaded_outputs)
    _dispose_widget(module, qapp)


def test_prepare_co_analysis_prompts_for_existing_word_output_path(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test prepare co analysis prompts for existing word output path.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        qapp: Parameter value (QApplication).
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(
        [
            str(tmp_path / "source.xlsx"),
            str(tmp_path / "co_description.xlsx"),
        ],
        emit_drop=True,
    )
    default_word = tmp_path / "CSE101_2025-26_CO_Analysis_Report.docx"
    default_word.write_text("existing", encoding="utf-8")
    replacement_word = tmp_path / "CSE101_2025-26_CO_Analysis_Report_v2.docx"
    captured: dict[str, object] = {}

    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getExistingDirectory",
        lambda *_args, **_kwargs: str(tmp_path),
    )

    def _fake_get_save_file_name(*_args, **_kwargs):
        return (str(replacement_word), "Word Files (*.docx)")

    monkeypatch.setattr(co_analysis_ui.QFileDialog, "getSaveFileName", _fake_get_save_file_name)
    monkeypatch.setattr(
        co_analysis_ui,
        "extract_course_metadata_and_students_from_workbook_path",
        lambda _path: (
            set(),
            {
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_COURSE_CODE_KEY): "CSE101",
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_ACADEMIC_YEAR_KEY): "2025-26",
            },
        ),
    )

    def _fake_generate_workbook(**kwargs):
        captured.update(kwargs)
        return type("_Result", (), {"output_path": str(tmp_path / "CSE101_2025-26_CO_Analysis.xlsx")})()

    monkeypatch.setattr(co_analysis_ui, "generate_workbook", _fake_generate_workbook)
    module._prepare_co_analysis_async()
    qapp.processEvents()

    context = cast(dict[str, object], captured.get("context", {}))
    assert str(context.get("co_description_path", "")).endswith("co_description.xlsx")
    assert str(context.get("word_output_path", "")) == str(replacement_word)
    _dispose_widget(module, qapp)


def test_submit_blocked_when_word_report_enabled_without_co_description(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test submit is blocked when report toggle is on and CO-description file is missing."""
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files([str(tmp_path / "source.xlsx")], emit_drop=True)
    called: dict[str, object] = {"started": False}
    monkeypatch.setattr(module, "_prepare_co_analysis_async", lambda: called.update(started=True))

    module.generate_word_report_checkbox.setChecked(True)
    module._on_submit_requested()
    qapp.processEvents()

    assert called["started"] is False
    _dispose_widget(module, qapp)


def test_submit_blocked_when_multiple_co_description_files_uploaded(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test submit is blocked when more than one CO-description file is uploaded with report toggle on."""
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(
        [
            str(tmp_path / "source.xlsx"),
            str(tmp_path / "co_description_a.xlsx"),
            str(tmp_path / "co_description_b.xlsx"),
        ],
        emit_drop=True,
    )
    called: dict[str, object] = {"started": False}
    monkeypatch.setattr(module, "_prepare_co_analysis_async", lambda: called.update(started=True))

    module.generate_word_report_checkbox.setChecked(True)
    module._on_submit_requested()
    qapp.processEvents()

    assert called["started"] is False
    _dispose_widget(module, qapp)


def test_submit_allows_marks_only_when_word_report_disabled(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test submit proceeds with marks-only upload when report toggle is off."""
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files([str(tmp_path / "source.xlsx")], emit_drop=True)
    called: dict[str, object] = {"started": False}
    monkeypatch.setattr(module, "_prepare_co_analysis_async", lambda: called.update(started=True))

    module.generate_word_report_checkbox.setChecked(False)
    module._on_submit_requested()
    qapp.processEvents()

    assert called["started"] is True
    _dispose_widget(module, qapp)


def test_prepare_co_analysis_with_word_toggle_off_emits_summary_toast(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    """Test generation feedback includes toast when word-report toggle is disabled."""
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files([str(tmp_path / "source.xlsx")], emit_drop=True)
    module.generate_word_report_checkbox.setChecked(False)
    captured_feedback: dict[str, object] = {}

    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getExistingDirectory",
        lambda *_args, **_kwargs: str(tmp_path),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "extract_course_metadata_and_students_from_workbook_path",
        lambda _path: (
            set(),
            {
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_COURSE_CODE_KEY): "CSE101",
                co_analysis_ui.normalize(co_analysis_ui.COURSE_METADATA_ACADEMIC_YEAR_KEY): "2025-26",
            },
        ),
    )
    monkeypatch.setattr(
        co_analysis_ui,
        "generate_workbook",
        lambda **_kwargs: type(
            "_Result",
            (),
            {"output_path": str(tmp_path / "CSE101_2025-26_CO_Analysis.xlsx")},
        )(),
    )

    def _capture_emit_workbook_generation_feedback(
        *,
        success_count: int,
        failed_count: int,
        channels: tuple[NotificationChannel, ...] = ("status", "activity_log", "toast"),
    ) -> None:
        captured_feedback["success_count"] = success_count
        captured_feedback["failed_count"] = failed_count
        captured_feedback["channels"] = channels

    monkeypatch.setattr(
        module._runtime,
        "emit_workbook_generation_feedback",
        _capture_emit_workbook_generation_feedback,
    )

    module._prepare_co_analysis_async()
    qapp.processEvents()

    assert captured_feedback["success_count"] == 1
    assert captured_feedback["failed_count"] == 0
    assert cast(tuple[NotificationChannel, ...], captured_feedback["channels"]) == (
        "status",
        "activity_log",
        "toast",
    )
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
