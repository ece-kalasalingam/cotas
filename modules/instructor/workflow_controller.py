"""Workflow-step UI state helpers for InstructorModule."""

from __future__ import annotations

from common.texts import t

_STEP_RUN_ALWAYS = (1, 2, 3)
_EMPTY_REASON = ""
_STEP_LIST_SEPARATOR = ". "
_STEP_LIST_STATE_GAP = "  "


class InstructorWorkflowController:
    def __init__(self, module: object) -> None:
        self._m = module

    def step_path(self, step: int) -> str | None:
        return getattr(self._m, self._m.PATH_ATTRS[step])

    def step_done(self, step: int) -> bool:
        return bool(getattr(self._m, self._m.DONE_ATTRS[step]))

    def step_outdated(self, step: int) -> bool:
        outdated_attr = self._m.OUTDATED_ATTRS.get(step)
        return bool(getattr(self._m, outdated_attr)) if outdated_attr else False

    def step_state_text(self, step: int) -> str:
        done = self.step_done(step)
        outdated = self.step_outdated(step)
        if done and outdated:
            return t("instructor.badge.needs_update")
        return t("instructor.badge.done") if done else t("instructor.badge.pending")

    def step_list_text(self, step: int) -> str:
        title = t(self._m.STEP_TITLE_KEYS[step])
        state = self.step_state_text(step)
        return f"{step}{_STEP_LIST_SEPARATOR}{title}{_STEP_LIST_STATE_GAP}{state}"

    def action_text_for_step(self, step: int) -> str:
        key = self._m.ACTION_REDO_KEYS[step] if self.step_done(step) else self._m.ACTION_DEFAULT_KEYS[step]
        return t(key)

    def can_run_step(self, step: int) -> tuple[bool, str]:
        if step in _STEP_RUN_ALWAYS:
            return True, _EMPTY_REASON
        return True, _EMPTY_REASON

    def on_step_selected(self, step: int) -> None:
        self._m.current_step = step
        self._m.state.current_step = step
        self._m._refresh_ui()
