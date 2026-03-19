"""Coordinator module message and toast helpers."""

from __future__ import annotations

from typing import Callable, cast

from PySide6.QtWidgets import QWidget

from common.texts import t
from common.toast import show_toast


def show_threshold_validation_toast(
    widget: object,
    *,
    message_key: str,
    title_key: str,
    toast_fn: Callable[..., None] = show_toast,
    translate: Callable[..., str] = t,
) -> None:
    toast_fn(
        cast(QWidget | None, widget),
        translate(message_key),
        title=translate(title_key),
        level="error",
    )
