"""Shared SheetSchema column-key helpers for COURSE_SETUP_V2 implementation."""

from __future__ import annotations

from common.error_catalog import validation_error_from_key
from common.sheet_schema import SheetSchema
from common.utils import normalize


def column_keys(sheet_schema: SheetSchema) -> tuple[str, ...]:
    raw = sheet_schema.sheet_rules.get("column_keys")
    if not isinstance(raw, (list, tuple)):
        return tuple()
    return tuple(normalize(value) for value in raw if isinstance(value, str) and normalize(value))


def column_index_by_key(sheet_schema: SheetSchema, key: str) -> int | None:
    wanted = normalize(key)
    for index, value in enumerate(column_keys(sheet_schema)):
        if value == wanted:
            return index
    return None


def required_column_index(sheet_schema: SheetSchema, key: str) -> int:
    index = column_index_by_key(sheet_schema, key)
    if index is not None:
        return index
    raise validation_error_from_key(
        "common.validation_failed_invalid_data",
        code="SCHEMA_COLUMN_KEY_MISSING",
        sheet_name=sheet_schema.name,
        column_key=key,
    )

