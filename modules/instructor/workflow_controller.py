"""Workflow-step UI state helpers for InstructorModule."""

from __future__ import annotations

from typing import Protocol, cast

from common.texts import t

_STEP_RUN_ALWAYS = (1, 2)
_EMPTY_REASON = ""


class InstructorWorkflowController:
    def __init__(self, module: object) -> None:
        self._m = cast(_InstructorWorkflowModule, module)

    def step_path(self, step: int) -> str | None:
        return getattr(self._m, self._m.PATH_ATTRS[step])

    def step_done(self, step: int) -> bool:
        return bool(getattr(self._m, self._m.DONE_ATTRS[step]))

    def step_outdated(self, step: int) -> bool:
        outdated_attr = self._m.OUTDATED_ATTRS.get(step)
        return bool(getattr(self._m, outdated_attr)) if outdated_attr else False

    def step_state_text(self, step: int) -> str:
        return ""

    def step_list_text(self, step: int) -> str:
        title = t(self._m.STEP_TITLE_KEYS[step])
        return title

    def action_text_for_step(self, step: int) -> str:
        key = self._m.ACTION_DEFAULT_KEYS[step]
        return t(key)

    def can_run_step(self, step: int) -> tuple[bool, str]:
        if step in _STEP_RUN_ALWAYS:
            return True, _EMPTY_REASON
        return True, _EMPTY_REASON

    def on_step_selected(self, step: int) -> None:
        self._m.current_step = step
        self._m.state.current_step = step
        self._m._refresh_ui()


class _InstructorState(Protocol):
    current_step: int


class _InstructorWorkflowModule(Protocol):
    PATH_ATTRS: dict[int, str]
    DONE_ATTRS: dict[int, str]
    OUTDATED_ATTRS: dict[int, str]
    STEP_TITLE_KEYS: dict[int, str]
    ACTION_DEFAULT_KEYS: dict[int, str]
    current_step: int
    state: _InstructorState

    def _refresh_ui(self) -> None:
        ...
