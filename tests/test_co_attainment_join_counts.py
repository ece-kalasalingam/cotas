from __future__ import annotations

import pytest

from domain.template_versions.course_setup_v2_impl import co_attainment


class _Sheet:
    def __init__(self, title: str) -> None:
        """Init.

        Args:
            title: Parameter value (str).

        Returns:
            None.

        Raises:
            None.
        """
        self.title = title


class _Workbook:
    def __init__(self) -> None:
        """Init.

        Args:
            None.

        Returns:
            None.

        Raises:
            None.
        """
        self.sheetnames = [
            co_attainment.co_direct_sheet_name(1),
            co_attainment.co_indirect_sheet_name(1),
        ]
        self._sheets = {name: _Sheet(name) for name in self.sheetnames}

    def __getitem__(self, key: str) -> _Sheet:
        """Getitem.

        Args:
            key: Parameter value (str).

        Returns:
            _Sheet: Return value.

        Raises:
            None.
        """
        return self._sheets[key]


def test_iter_co_rows_returns_counts_without_rescans(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test iter co rows returns counts without rescans.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    calls = {"direct": 0, "indirect": 0}
    direct_name = co_attainment.co_direct_sheet_name(1)

    direct_rows = [
        co_attainment._ParsedScoreRow(
            reg_hash=1,
            reg_key="r1",
            reg_no="R1",
            student_name="A",
            score=40.0,
        ),
        co_attainment._ParsedScoreRow(
            reg_hash=2,
            reg_key="r2",
            reg_no="R2",
            student_name="B",
            score=50.0,
        ),
    ]
    indirect_rows = [
        co_attainment._ParsedScoreRow(
            reg_hash=2,
            reg_key="r2",
            reg_no="R2",
            student_name="B",
            score=20.0,
        ),
        co_attainment._ParsedScoreRow(
            reg_hash=3,
            reg_key="r3",
            reg_no="R3",
            student_name="C",
            score=30.0,
        ),
    ]

    def _fake_iter_score_rows(sheet, *, ratio: float):  # noqa: ANN001
        """Fake iter score rows.

        Args:
            sheet: Parameter value.
            ratio: Parameter value (float).

        Returns:
            list[co_attainment._ParsedScoreRow]: Return value.

        Raises:
            None.
        """
        del ratio
        if sheet.title == direct_name:
            calls["direct"] += 1
            return iter(direct_rows)
        calls["indirect"] += 1
        return iter(indirect_rows)

    monkeypatch.setattr(co_attainment, "_iter_score_rows", _fake_iter_score_rows)

    rows, direct_total, indirect_total, dropped, direct_columns, indirect_columns = co_attainment._iter_co_rows_from_workbook(
        _Workbook(),
        co_index=1,
        workbook_name="w.xlsx",
    )

    assert calls == {"direct": 1, "indirect": 1}
    assert direct_total == 2
    assert indirect_total == 2
    assert dropped == 2
    assert len(rows) == 1
    assert rows[0].reg_no == "R2"
    assert direct_columns == []
    assert indirect_columns == []


def test_direct_total_100_treats_absent_as_zero() -> None:
    """Test direct total conversion treats absent/NA tokens as zero.

    Args:
        None.

    Returns:
        None.

    Raises:
        None.
    """
    assert co_attainment._direct_total_100_from_direct_score(co_attainment.CO_REPORT_ABSENT_TOKEN) == 0.0
    assert co_attainment._direct_total_100_from_direct_score(co_attainment.CO_REPORT_NOT_APPLICABLE_TOKEN) == 0.0


def test_co_direct_total_100_uses_zero_only_for_absent_assessments() -> None:
    """Test pass-sheet CO total keeps non-absent marks and maps absent assessments to zero.

    Args:
        None.

    Returns:
        None.

    Raises:
        None.
    """
    row = co_attainment._CoAttainmentRow(
        reg_hash=10,
        reg_no="R10",
        student_name="S",
        direct_score=0.0,
        indirect_score=0.0,
        worksheet_name="CO1_Direct",
        workbook_name="w.xlsx",
    )
    columns = {
        1: [
            co_attainment._DirectComponentColumn(name="CAT", max_marks=10.0, weight=50.0, score_column=4),
            co_attainment._DirectComponentColumn(name="SEE", max_marks=10.0, weight=50.0, score_column=6),
        ]
    }
    scores = {
        1: {
            10: {
                "cat": co_attainment.CO_REPORT_ABSENT_TOKEN,
                "see": 8.0,
            }
        }
    }
    value = co_attainment._co_direct_total_100_for_row(
        co_index=1,
        row=row,
        direct_columns_by_co=columns,
        direct_scores_by_co=scores,
    )
    assert value == 40.0
