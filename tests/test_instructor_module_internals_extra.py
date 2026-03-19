from __future__ import annotations

from datetime import datetime
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from common.jobs import CancellationToken
from modules import instructor_module as instructor_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module(monkeypatch: pytest.MonkeyPatch) -> instructor_ui.InstructorModule:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    return instructor_ui.InstructorModule()


def test_top_level_compat_wrappers_delegate(monkeypatch: pytest.MonkeyPatch) -> None:
    seen: list[tuple[str, object]] = []

    monkeypatch.setattr(
        instructor_ui,
        "_publish_status_compat_impl",
        lambda **kwargs: seen.append(("publish", kwargs["message"])),
    )
    monkeypatch.setattr(instructor_ui, "build_i18n_log_message", lambda key, fallback="": f"P:{key}:{fallback}")
    monkeypatch.setattr(instructor_ui, "_set_busy_compat_impl", lambda **kwargs: seen.append(("busy", kwargs["busy"])))

    target = type("_T", (), {"_supports_i18n_status_payload": True})()
    msg = instructor_ui.t("instructor.status.operation_cancelled")
    instructor_ui._publish_status_compat(target, msg)
    instructor_ui._set_busy_compat(target, True, job_id="j")

    monkeypatch.setattr(instructor_ui, "validate_filled_marks_manifest_schema_by_template", lambda *_a, **_k: seen.append(("validate", True)))
    monkeypatch.setattr(instructor_ui, "filled_marks_manifest_validators", lambda: {"x": object()})
    instructor_ui._validate_filled_marks_manifest_schema_by_template(object(), {}, template_id="x")
    assert instructor_ui._filled_marks_manifest_validators().keys() == {"x"}

    assert any(tag == "publish" and str(value).startswith("P:instructor.status.operation_cancelled") for tag, value in seen)
    assert ("busy", True) in seen
    assert ("validate", True) in seen


def test_refresh_ui_step_specific_paths_and_busy_disable(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    monkeypatch.setattr(module, "_refresh_quick_links", lambda: None)

    module.current_step = 1
    module.step2_upload_ready = False
    monkeypatch.setattr(module, "_can_run_step", lambda _s: (False, "reason-blocked"))
    module._refresh_ui()
    assert module.step1_upload_action.text() == "instructor.action.step1.upload"
    assert module.step1_prepare_action.text() == "instructor.action.step1.prepare"
    assert module.active_note.text() == "reason-blocked"

    module.current_step = 2
    module.filled_marks_done = True
    module.filled_marks_outdated = False
    module.final_report_done = True
    monkeypatch.setattr(module, "_can_run_step", lambda _s: (True, ""))
    module._refresh_ui()
    assert module.step2_upload_action.text() == "instructor.action.step2.upload.default"
    assert module.step2_generate_action.text() == "instructor.action.step2.generate.default"
    assert module.step2_generate_action.isEnabled() is False

    module.state.busy = True
    module._refresh_ui()
    assert module.primary_action.isEnabled() is False
    assert module.step1_upload_action.isEnabled() is False
    assert module.step1_prepare_action.isEnabled() is False
    assert module.step2_upload_action.isEnabled() is False
    assert module.step2_generate_action.isEnabled() is False
    module.close()


def test_user_log_append_publish_and_rerender_branches(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    monkeypatch.setattr(instructor_ui, "parse_i18n_log_message", lambda _m: None)
    monkeypatch.setattr(instructor_ui, "resolve_i18n_log_message", lambda m: m)
    monkeypatch.setattr(instructor_ui, "format_log_line_at", lambda message, timestamp=None: f"L:{message}")
    module._append_user_log("plain-msg")
    assert module._user_log_entries and module._user_log_entries[-1]["message"] == "plain-msg"

    monkeypatch.setattr(instructor_ui, "parse_i18n_log_message", lambda _m: ("k", {"n": 1}, "fb"))
    monkeypatch.setattr(instructor_ui, "resolve_i18n_log_message", lambda _m: "localized")
    monkeypatch.setattr(instructor_ui, "format_log_line_at", lambda *_a, **_k: None)
    module._append_user_log("i18n-msg")
    assert module._user_log_entries[-1]["text_key"] == "k"

    emitted: list[str] = []
    monkeypatch.setattr(module, "_append_user_log", lambda message: emitted.append(f"append:{message}"))
    monkeypatch.setattr(instructor_ui, "emit_user_status", lambda _sig, msg, logger=None: emitted.append(f"emit:{msg}"))
    module._publish_status("status-1")
    assert emitted == ["append:status-1", "emit:status-1"]

    module._user_log_entries = [
        {
            "timestamp": datetime.now(),
            "text_key": "k.fail",
            "kwargs": {"x": 1},
            "fallback": "fallback-text",
            "message": "m1",
        },
        {
            "timestamp": datetime.now(),
            "message": "m2",
        },
    ]

    def _t(key: str, **kwargs):
        if key == "k.fail":
            raise RuntimeError("boom")
        return f"T:{key}:{kwargs}"

    monkeypatch.setattr(instructor_ui, "t", _t)
    lines: list[str] = []
    monkeypatch.setattr(module.user_log_view, "appendPlainText", lambda line: lines.append(line))
    monkeypatch.setattr(module.user_log_view, "clear", lambda: lines.clear())
    monkeypatch.setattr(instructor_ui, "format_log_line_at", lambda message, timestamp=None: None if message == "m2" else f"R:{message}")

    module._rerender_user_log()
    assert lines == ["R:fallback-text"]
    module.close()


def test_misc_wrappers_shortcuts_and_close_cleanup(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    calls = {"u2": 0, "p2": 0, "u3": 0, "g3": 0, "refresh": 0}
    monkeypatch.setattr(module, "_upload_course_details_async", lambda: calls.__setitem__("u2", calls["u2"] + 1))
    monkeypatch.setattr(module, "_prepare_marks_template_async", lambda: calls.__setitem__("p2", calls["p2"] + 1))
    monkeypatch.setattr(module, "_upload_filled_marks_async", lambda: calls.__setitem__("u3", calls["u3"] + 1))
    monkeypatch.setattr(module, "_generate_final_report_async", lambda: calls.__setitem__("g3", calls["g3"] + 1))
    monkeypatch.setattr(module, "_refresh_ui", lambda: calls.__setitem__("refresh", calls["refresh"] + 1))

    module._on_step1_upload_clicked()
    module._on_step1_prepare_clicked()
    module._on_step2_upload_clicked()
    module._on_step2_generate_clicked()
    assert calls == {"u2": 1, "p2": 1, "u3": 0, "g3": 1, "refresh": 4}

    module.current_step = 1
    module.step1_upload_action.setEnabled(True)
    module.state.busy = False
    module._on_open_shortcut_activated()
    assert calls["u2"] == 2

    module.current_step = 2
    module.step2_upload_action.setEnabled(True)
    module._on_open_shortcut_activated()
    assert calls["u3"] == 0

    module.current_step = 1
    module.step1_prepare_action.setEnabled(True)
    module._on_save_shortcut_activated()

    module.current_step = 2
    module.step2_generate_action.setEnabled(True)
    module._on_save_shortcut_activated()

    assert calls["p2"] >= 1
    assert calls["g3"] >= 1

    monkeypatch.setattr(module, "_quick_links_html", lambda: "outputs")
    module.set_shared_activity_log_mode(True)
    assert module.info_tabs.isHidden() is True
    module.set_shared_activity_log_mode(False)
    assert module.info_tabs.isHidden() is False
    assert module.get_shared_outputs_html() == "outputs"

    # _on_quick_link_activated empty path early-return
    monkeypatch.setattr(instructor_ui.QDesktopServices, "openUrl", lambda _url: (_ for _ in ()).throw(AssertionError("should not open")))
    module._on_quick_link_activated("file::   ")

    # _start_async_operation wrapper
    seen_start: list[str] = []

    class _Runner:
        def start(self, **kwargs):
            seen_start.append(kwargs.get("job_id") or "")

    module._async_runner = cast(Any, _Runner())
    module._start_async_operation(
        token=CancellationToken(),
        job_id="job-x",
        work=lambda: None,
        on_success=lambda _r: None,
        on_failure=lambda _e: None,
    )
    assert seen_start == ["job-x"]

    removed = {"count": 0}
    monkeypatch.setattr(instructor_ui._logger, "removeHandler", lambda _h: removed.__setitem__("count", removed["count"] + 1))
    module._ui_log_handler = cast(Any, object())
    module.close()
    assert removed["count"] == 1


def test_setup_ui_logging_and_step_toast_wrappers(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    original_setup = instructor_ui.InstructorModule._setup_ui_logging
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    monkeypatch.setattr(instructor_ui, "t", lambda key, **kwargs: key)
    module = instructor_ui.InstructorModule()

    seen = {"handler": 0, "append": []}

    class _Handler:
        def __init__(self, callback):
            self.callback = callback

    monkeypatch.setattr(instructor_ui, "UILogHandler", _Handler)
    monkeypatch.setattr(instructor_ui._logger, "addHandler", lambda _h: seen.__setitem__("handler", seen["handler"] + 1))
    monkeypatch.setattr(module, "_append_user_log", lambda message: seen["append"].append(message))
    monkeypatch.setattr(instructor_ui, "build_i18n_log_message", lambda key, fallback="": f"I18N:{key}:{fallback}")

    original_setup(module)
    assert seen["handler"] == 1
    assert seen["append"] and seen["append"][0].startswith("I18N:instructor.log.ready")
    # Early-return branch when handler already initialized.
    original_setup(module)
    assert seen["handler"] == 1

    toast_calls: list[tuple[str, object]] = []
    monkeypatch.setattr(instructor_ui, "show_step_success_toast", lambda _m, **kwargs: toast_calls.append(("success", kwargs)))
    monkeypatch.setattr(instructor_ui, "show_validation_error_toast", lambda _m, msg: toast_calls.append(("validation", msg)))
    monkeypatch.setattr(instructor_ui, "show_system_error_toast", lambda _m, **kwargs: toast_calls.append(("system", kwargs)))

    module._show_step_success_toast(1)
    module._show_validation_error_toast("bad")
    module._show_system_error_toast(1)

    assert toast_calls[0][0] == "success"
    assert toast_calls[1] == ("validation", "bad")
    assert toast_calls[2][0] == "system"
    module.close()


def test_additional_wrapper_and_ui_branches(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    monkeypatch.setattr(instructor_ui, "build_final_report_default_name", lambda _p: "final.xlsx")
    monkeypatch.setattr(instructor_ui, "_shared_atomic_copy_file", lambda *_a, **_k: "copied")
    assert instructor_ui._build_final_report_default_name("x.xlsx") == "final.xlsx"
    assert instructor_ui._atomic_copy_file("a", "b") == "copied"

    module = _build_module(monkeypatch)
    monkeypatch.setattr(module, "_refresh_quick_links", lambda: None)
    assert module._quick_links_html() is not None
    assert module._step_path(1) == module._workflow_controller.step_path(1)
    assert module._step_done(1) == module._workflow_controller.step_done(1)
    assert module._step_outdated(2) == module._workflow_controller.step_outdated(2)
    assert module._step_state_text(1) == module._workflow_controller.step_state_text(1)
    assert module._action_text_for_step(1) == module._workflow_controller.action_text_for_step(1)

    selected = {"step": 0}
    monkeypatch.setattr(module._workflow_controller, "on_step_selected", lambda step: selected.__setitem__("step", step))
    module._on_step_selected(2)
    assert selected["step"] == 2

    # Cover selection clearing branch.
    module.user_log_view.setPlainText("one")
    module.generated_outputs_view.setPlainText("two")
    c1 = module.user_log_view.textCursor()
    c1.select(c1.SelectionType.Document)
    module.user_log_view.setTextCursor(c1)
    c2 = module.generated_outputs_view.textCursor()
    c2.select(c2.SelectionType.Document)
    module.generated_outputs_view.setTextCursor(c2)
    module._clear_info_text_selection()
    assert module.user_log_view.textCursor().hasSelection() is False
    assert module.generated_outputs_view.textCursor().hasSelection() is False

    # Invalid step branch resets to first workflow step.
    module.current_step = 99
    module._refresh_ui()
    assert module.current_step == 1

    # Outdated-note branches.
    module.current_step = 2
    module.filled_marks_outdated = True
    module.final_report_outdated = False
    monkeypatch.setattr(module, "_can_run_step", lambda _s: (True, ""))
    monkeypatch.setattr(module, "_step_outdated", lambda _s: True)
    module._refresh_ui()
    assert module.active_note.text() == "instructor.note.outdated_current"
    monkeypatch.setattr(module, "_step_outdated", lambda _s: False)
    module._refresh_ui()
    assert module.active_note.text() == "instructor.note.outdated_downstream"

    # Retranslate wrapper.
    called = {"rerender": 0, "refresh": 0, "clear": 0}
    monkeypatch.setattr(module, "_rerender_user_log", lambda: called.__setitem__("rerender", called["rerender"] + 1))
    monkeypatch.setattr(module, "_refresh_ui", lambda: called.__setitem__("refresh", called["refresh"] + 1))
    monkeypatch.setattr(module, "_clear_info_text_selection", lambda: called.__setitem__("clear", called["clear"] + 1))
    module.retranslate_ui()
    assert called == {"rerender": 1, "refresh": 1, "clear": 1}

    # Download course template click busy/idle branches.
    called_dl = {"count": 0, "refresh": 0}
    monkeypatch.setattr(module, "_download_course_template_async", lambda: called_dl.__setitem__("count", called_dl["count"] + 1))
    monkeypatch.setattr(module, "_refresh_ui", lambda: called_dl.__setitem__("refresh", called_dl["refresh"] + 1))
    module.state.busy = True
    module._on_download_course_template_clicked()
    module.state.busy = False
    module._on_download_course_template_clicked()
    assert called_dl["count"] == 1
    assert called_dl["refresh"] == 1

    # _setup_ui_logging early return branch.
    module._ui_log_handler = cast(Any, object())
    module._setup_ui_logging()

    # Direct async wrappers and legacy wrappers.
    direct_calls = {"prepare": 0, "download": 0, "generate": 0}
    monkeypatch.setattr(instructor_ui, "prepare_marks_template_async", lambda _m, ns=None: direct_calls.__setitem__("prepare", direct_calls["prepare"] + 1))
    monkeypatch.setattr(instructor_ui, "download_course_template_async", lambda _m, ns=None: direct_calls.__setitem__("download", direct_calls["download"] + 1))
    monkeypatch.setattr(instructor_ui, "generate_final_report_async", lambda _m, ns=None: direct_calls.__setitem__("generate", direct_calls["generate"] + 1))
    module._prepare_marks_template_async()
    module._download_course_template_async()
    module._generate_final_report_async()
    module._download_course_template()
    module._prepare_marks_template()
    module._generate_final_report()
    assert direct_calls["prepare"] == 2
    assert direct_calls["download"] >= 1
    assert direct_calls["generate"] == 2
    module.close()

