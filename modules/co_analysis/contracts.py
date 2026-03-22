"""Runtime contract checks for CO Analysis integration namespaces."""

from __future__ import annotations

from collections.abc import Mapping


def require_keys(namespace: Mapping[str, object], *, keys: tuple[str, ...], context: str) -> None:
    missing = [key for key in keys if key not in namespace]
    if missing:
        ordered = ", ".join(sorted(missing))
        raise RuntimeError(f"{context} namespace is missing required keys: {ordered}")

