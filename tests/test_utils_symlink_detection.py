from __future__ import annotations

import pytest

from common.exceptions import ValidationError
from common import utils


def test_assert_not_symlink_path_allows_regular_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    """Assert not symlink path allows regular paths.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    monkeypatch.setattr(utils, "path_uses_symlink", lambda _path: False)
    utils.assert_not_symlink_path("C:/normal.xlsx")


def test_assert_not_symlink_path_raises_for_symlink_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    """Assert not symlink path raises for symlink paths.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    monkeypatch.setattr(utils, "path_uses_symlink", lambda _path: True)

    with pytest.raises(ValidationError) as excinfo:
        utils.assert_not_symlink_path("C:/linked.xlsx")

    assert excinfo.value.code == "WORKBOOK_SYMLINK_NOT_ALLOWED"
