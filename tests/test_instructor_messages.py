from __future__ import annotations

from modules.instructor import messages


def test_localized_log_messages_builds_success_and_error_messages(monkeypatch) -> None:
    monkeypatch.setattr(messages, "t", lambda key, **kwargs: f"T:{key}:{kwargs.get('process', '')}")
    monkeypatch.setattr(
        messages,
        "build_i18n_log_message",
        lambda key, kwargs=None, fallback="": f"L:{key}:{kwargs}:{fallback}",
    )

    success, error = messages.localized_log_messages("instructor.process.step1")

    assert "instructor.log.completed_process" in success
    assert "instructor.log.error_while_process" in error
    assert "instructor.process.step1" in success
    assert "instructor.process.step1" in error


def test_show_step_success_toast_uses_localized_title_and_step(monkeypatch) -> None:
    seen: list[tuple[object, str, str, str]] = []
    monkeypatch.setattr(messages, "t", lambda key, **kwargs: f"T:{key}:{kwargs}")
    monkeypatch.setattr(
        messages,
        "show_toast",
        lambda widget, body, *, title, level: seen.append((widget, body, title, level)),
    )

    obj = object()
    messages.show_step_success_toast(obj, step=2, title_key="instructor.step2.title")

    assert len(seen) == 1
    widget, body, title, level = seen[0]
    assert widget is obj
    assert "step_completed" in body
    assert "step': 2" in body
    assert "success_title" in title
    assert level == "success"


def test_show_validation_error_toast_uses_error_title(monkeypatch) -> None:
    seen: list[tuple[str, str, str]] = []
    monkeypatch.setattr(messages, "t", lambda key, **kwargs: f"T:{key}")
    monkeypatch.setattr(
        messages,
        "show_toast",
        lambda _widget, body, *, title, level: seen.append((body, title, level)),
    )

    messages.show_validation_error_toast(object(), "Invalid row")

    assert seen == [("Invalid row", "T:instructor.msg.validation_title", "error")]


def test_show_system_error_toast_uses_action_translation(monkeypatch) -> None:
    seen: list[tuple[str, str, str]] = []

    def _t(key: str, **kwargs):
        return f"T:{key}:{kwargs.get('action', '')}"

    monkeypatch.setattr(messages, "t", _t)
    monkeypatch.setattr(
        messages,
        "show_toast",
        lambda _widget, body, *, title, level: seen.append((body, title, level)),
    )

    messages.show_system_error_toast(object(), title_key="instructor.step2.title")

    assert len(seen) == 1
    body, title, level = seen[0]
    assert "failed_to_do" in body
    assert "instructor.step2.title" in body
    assert "error_title" in title
    assert level == "error"
