"""Shared output markup helpers."""

from __future__ import annotations

from html import escape

LABEL_BOLD_TEMPLATE = "<b>{label}</b>: {value}"


def render_labeled_value(label: str, value: str) -> str:
    """Render a consistent 'Label: Value' fragment with bold label."""
    return LABEL_BOLD_TEMPLATE.format(label=escape(label), value=value)

