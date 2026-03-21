from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Callable, cast

import pytest

from common.exceptions import AppSystemError
from domain import instructor_template_engine as mod


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="Windows-only permission integration check")
def test_generate_course_template_reports_permission_failure_on_locked_destination(tmp_path: Path) -> None:
    pytest.importorskip("xlsxwriter")
    import msvcrt
    msvcrt_api = msvcrt if sys.platform.startswith("win") else None
    if msvcrt_api is None:
        pytest.skip("Windows-only msvcrt locking API")

    locking = getattr(msvcrt_api, "locking", None)
    lock_non_block = getattr(msvcrt_api, "LK_NBLCK", None)
    unlock = getattr(msvcrt_api, "LK_UNLCK", None)
    if locking is None or lock_non_block is None or unlock is None:
        pytest.skip("Windows msvcrt locking symbols unavailable")
    locking_fn = cast(Callable[[int, int, int], Any], locking)
    lock_non_block_value = cast(int, lock_non_block)
    unlock_value = cast(int, unlock)

    output = tmp_path / "course_setup.xlsx"
    output.write_bytes(b"locked")

    with output.open("r+b") as handle:
        handle.seek(0)
        locking_fn(handle.fileno(), lock_non_block_value, 1)
        try:
            with pytest.raises(AppSystemError):
                mod.generate_course_details_template(output)
        finally:
            handle.seek(0)
            locking_fn(handle.fileno(), unlock_value, 1)

    assert output.read_bytes() == b"locked"
