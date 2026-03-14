from __future__ import annotations

import sys
from pathlib import Path

import pytest

from common.exceptions import AppSystemError
from domain import instructor_template_engine as mod


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="Windows-only permission integration check")
def test_generate_course_template_reports_permission_failure_on_locked_destination(tmp_path: Path) -> None:
    pytest.importorskip("xlsxwriter")
    import msvcrt

    output = tmp_path / "course_setup.xlsx"
    output.write_bytes(b"locked")

    with output.open("r+b") as handle:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        try:
            with pytest.raises(AppSystemError):
                mod.generate_course_details_template(output)
        finally:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)

    assert output.read_bytes() == b"locked"
