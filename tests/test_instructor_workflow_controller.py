from dataclasses import dataclass

from modules.instructor.workflow_controller import InstructorWorkflowController


@dataclass
class _State:
    current_step: int = 1


class _Module:
    PATH_ATTRS = {1: "p1", 2: "p2"}
    DONE_ATTRS = {1: "d1", 2: "d2"}
    OUTDATED_ATTRS = {2: "o2"}
    STEP_TITLE_KEYS = {1: "k1", 2: "k2"}
    ACTION_DEFAULT_KEYS = {1: "a1", 2: "a2"}

    def __init__(self) -> None:
        self.p1 = "x"
        self.p2 = None
        self.d1 = False
        self.d2 = True
        self.o2 = True
        self.step2_upload_ready = False
        self.current_step = 1
        self.state = _State()
        self.refreshed = 0

    def _refresh_ui(self) -> None:
        self.refreshed += 1


def test_workflow_controller_remaining_branches(monkeypatch) -> None:
    from modules.instructor import workflow_controller as wc

    monkeypatch.setattr(wc, "t", lambda key: key)
    mod = _Module()
    controller = InstructorWorkflowController(mod)

    assert controller.step_path(1) == "x"
    assert controller.step_done(2) is True
    assert controller.step_outdated(2) is True
    assert controller.step_state_text(2) == ""
    assert controller.step_list_text(1) == "k1"
    assert controller.action_text_for_step(2) == "a2"
    assert controller.can_run_step(2) == (False, "instructor.require.step1")
    mod.step2_upload_ready = True
    assert controller.can_run_step(2) == (True, "")
    assert controller.can_run_step(99) == (True, "")

    controller.on_step_selected(2)
    assert mod.current_step == 2
    assert mod.state.current_step == 2
    assert mod.refreshed == 1
