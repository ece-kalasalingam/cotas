"""Shared CO token parser for validation and report-generation flows."""

from __future__ import annotations

import re
from typing import Any

_CO_TOKEN_SEPARATOR = chr(44)
_CO_TOKEN_PATTERN = "".join((r"(?:", "co", r")?\s*", r"(\d+)"))


def parse_co_tokens(value: Any, *, dedupe: bool = False) -> list[int]:
    """Parse co tokens.
    
    Args:
        value: Parameter value (Any).
        dedupe: Parameter value (bool).
    
    Returns:
        list[int]: Return value.
    
    Raises:
        None.
    """
    if value is None:
        return []
    if isinstance(value, bool):
        return []
    if isinstance(value, int):
        return [value]
    if isinstance(value, float):
        return [int(value)] if value.is_integer() else []

    token = str(value).strip()
    if not token:
        return []

    values: list[int] = []
    for item in token.split(_CO_TOKEN_SEPARATOR):
        part = item.strip()
        if not part:
            return []
        match = re.fullmatch(_CO_TOKEN_PATTERN, part, flags=re.IGNORECASE)
        if not match:
            return []
        values.append(int(match.group(1)))

    if not dedupe:
        return values
    seen: set[int] = set()
    out: list[int] = []
    for entry in values:
        if entry in seen:
            continue
        seen.add(entry)
        out.append(entry)
    return out
