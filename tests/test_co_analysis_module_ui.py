from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import co_analysis_module as co_analysis_ui


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


def _build_module_with_message_capture(
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[co_analysis_ui.COAnalysisModule, list[str]]:
    seen_keys: list[str] = []
    monkeypatch.setattr(co_analysis_ui, "t", lambda key, **kwargs: key)

    def _capture_notify_message_key(self, text_key: str, **kwargs: object) -> None:
        del self, kwargs
        seen_keys.append(text_key)

    monkeypatch.setattr(
        co_analysis_ui.ModuleRuntime,
        "notify_message_key",
        _capture_notify_message_key,
    )
    monkeypatch.setattr(co_analysis_ui, "validate_uploaded_source_workbook", lambda _path: None)
    monkeypatch.setattr(co_analysis_ui, "consume_last_source_anomaly_warnings", lambda: [])
    monkeypatch.setattr(co_analysis_ui.COAnalysisModule, "_setup_ui_logging", lambda self: None)
    return co_analysis_ui.COAnalysisModule(), seen_keys


def test_drop_widget_syncs_files_and_clear(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
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


def test_submit_reports_no_backend_action(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module._on_submit_requested()
    assert "coordinator.status.operation_cancelled" in seen_keys
    _dispose_widget(module, qapp)


def test_submit_shows_threshold_violation_when_invalid(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
) -> None:
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module.threshold_l1_input.setValue(70.0)
    module.threshold_l2_input.setValue(60.0)
    module.threshold_l3_input.setValue(90.0)
    module._on_submit_requested()
    assert "coordinator.thresholds.invalid_rule" in seen_keys
    _dispose_widget(module, qapp)


def test_download_co_description_template_link_triggers_generation(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
    tmp_path: Path,
) -> None:
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    output_path = tmp_path / "CO_Description_Template.xlsx"
    called: dict[str, object] = {}

    monkeypatch.setattr(
        co_analysis_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: (str(output_path), "Excel Files (*.xlsx)"),
    )

    def _fake_generate_workbook(**kwargs):
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
