from __future__ import annotations

import pytest

from common import runtime_dependency_guard as dep_guard
from common.exceptions import ConfigurationError


def test_missing_runtime_dependency_packages_reports_only_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test missing runtime dependency packages reports only missing.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """

    def _fake_find_spec(name: str):
        if name == "docx":
            return None
        return object()

    monkeypatch.setattr(dep_guard, "find_spec", _fake_find_spec)
    missing = dep_guard.missing_runtime_dependency_packages()
    if not (missing == ("python-docx",)):
        raise AssertionError('assertion failed')


def test_runtime_dependency_spec_rejects_unknown() -> None:
    """Test runtime dependency spec rejects unknown.

    Args:
        None.

    Returns:
        None.

    Raises:
        None.
    """
    with pytest.raises(ConfigurationError):
        _ = dep_guard.runtime_dependency_spec("unknown-lib")
