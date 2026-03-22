"""Workflow state helpers for COAnalysisModule."""

from __future__ import annotations

from typing import Protocol, cast

from modules.co_analysis.validators.attainment_thresholds import (
    has_valid_attainment_thresholds,
    has_valid_co_attainment_percent,
)


class COAnalysisWorkflowController:
    def __init__(self, module: object) -> None:
        self._m = cast(_COAnalysisWorkflowModule, module)

    def read_attainment_thresholds(self) -> tuple[float, float, float]:
        return (
            float(self._m.threshold_l1_input.value()),
            float(self._m.threshold_l2_input.value()),
            float(self._m.threshold_l3_input.value()),
        )

    def has_valid_attainment_thresholds(self) -> bool:
        l1, l2, l3 = self.read_attainment_thresholds()
        return has_valid_attainment_thresholds(l1, l2, l3)

    def read_co_attainment_target(self) -> tuple[float, int]:
        raw_level = self._m.co_attainment_level_input.currentData()
        level = int(raw_level) if isinstance(raw_level, int) else 1
        return (
            float(self._m.co_attainment_percent_input.value()),
            level,
        )

    def has_valid_co_attainment_target(self) -> bool:
        percent, level = self.read_co_attainment_target()
        return has_valid_co_attainment_percent(percent) and level >= 1

    def notify_threshold_violation(self, *, force: bool) -> None:
        if self._m._threshold_violation_active and not force:
            return
        if not self.has_valid_attainment_thresholds():
            key = self._m._THRESHOLD_VALIDATION_KEY
        else:
            key = self._m._CO_ATTAINMENT_TARGET_VALIDATION_KEY
        self._m._show_threshold_validation_toast(message_key=key)
        self._m._publish_status_key(key)
        self._m._threshold_violation_active = True

    def on_threshold_value_changed(self) -> None:
        if self.has_valid_attainment_thresholds() and self.has_valid_co_attainment_target():
            self._m._threshold_violation_active = False

    def on_threshold_editing_finished(self) -> None:
        if self.has_valid_attainment_thresholds() and self.has_valid_co_attainment_target():
            self._m._threshold_violation_active = False
            return
        self.notify_threshold_violation(force=False)


class _ThresholdInput(Protocol):
    def value(self) -> float:
        ...


class _LevelInput(Protocol):
    def currentData(self) -> object:
        ...


class _COAnalysisWorkflowModule(Protocol):
    _THRESHOLD_VALIDATION_KEY: str
    _CO_ATTAINMENT_TARGET_VALIDATION_KEY: str
    _threshold_violation_active: bool
    threshold_l1_input: _ThresholdInput
    threshold_l2_input: _ThresholdInput
    threshold_l3_input: _ThresholdInput
    co_attainment_percent_input: _ThresholdInput
    co_attainment_level_input: _LevelInput

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        ...

    def _show_threshold_validation_toast(self, *, message_key: str) -> None:
        ...

