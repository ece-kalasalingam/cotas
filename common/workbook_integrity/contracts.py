"""Typed contracts for workbook integrity payloads."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class SystemWorkbookPayload:
    template_id: str
    template_hash: str
    manifest: dict[str, Any]

