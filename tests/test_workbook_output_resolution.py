from __future__ import annotations

from common.workbook_output_resolution import (
    OverwriteConflict,
    extract_overwrite_conflicts_from_generation_result,
    resolve_overwrite_conflicts,
)


def test_extract_overwrite_conflicts_from_generation_result_filters_expected_items() -> None:
    result = {
        "results": {
            "a": {
                "status": "failed",
                "reason": "output_already_exists",
                "source_path": "C:/src/a.xlsx",
                "existing_output_path": "C:/out/a_marks.xlsx",
            },
            "b": {
                "status": "failed",
                "reason": "other_error",
                "source_path": "C:/src/b.xlsx",
                "output_path": "C:/out/b_marks.xlsx",
            },
            "c": {
                "status": "generated",
                "reason": None,
                "source_path": "C:/src/c.xlsx",
                "output_path": "C:/out/c_marks.xlsx",
            },
            "d": {
                "status": "failed",
                "reason": "output_already_exists",
                "source_path": "C:/src/d.xlsx",
                "output_path": "C:/out/d_marks.xlsx",
            },
        }
    }

    conflicts = extract_overwrite_conflicts_from_generation_result(result)

    assert [(item.source_path, item.output_path) for item in conflicts] == [
        ("C:/src/a.xlsx", "C:/out/a_marks.xlsx"),
        ("C:/src/d.xlsx", "C:/out/d_marks.xlsx"),
    ]


def test_resolve_overwrite_conflicts_uses_bulk_confirmation_when_above_limit() -> None:
    conflicts = [
        OverwriteConflict(source_path="C:/src/a.xlsx", output_path="C:/out/a.xlsx"),
        OverwriteConflict(source_path="C:/src/b.xlsx", output_path="C:/out/b.xlsx"),
        OverwriteConflict(source_path="C:/src/c.xlsx", output_path="C:/out/c.xlsx"),
    ]
    observed_paths: list[list[str]] = []

    def _ask_overwrite_all(paths: list[str]) -> bool:
        observed_paths.append(paths)
        return True

    def _ask_output_path(_suggested_output: str) -> str | None:
        raise AssertionError("Per-file resolver must not be called in bulk mode.")

    resolved = resolve_overwrite_conflicts(
        conflicts,
        per_file_native_limit=2,
        ask_overwrite_all=_ask_overwrite_all,
        ask_output_path=_ask_output_path,
    )

    assert observed_paths == [["C:/out/a.xlsx", "C:/out/b.xlsx", "C:/out/c.xlsx"]]
    assert resolved.retry_sources == ["C:/src/a.xlsx", "C:/src/b.xlsx", "C:/src/c.xlsx"]
    assert resolved.output_path_overrides == {
        "C:/src/a.xlsx": "C:/out/a.xlsx",
        "C:/src/b.xlsx": "C:/out/b.xlsx",
        "C:/src/c.xlsx": "C:/out/c.xlsx",
    }


def test_resolve_overwrite_conflicts_uses_per_file_resolution_when_within_limit() -> None:
    conflicts = [
        OverwriteConflict(source_path="C:/src/a.xlsx", output_path="C:/out/a.xlsx"),
        OverwriteConflict(source_path="C:/src/b.xlsx", output_path="C:/out/b.xlsx"),
    ]
    suggested_outputs: list[str] = []

    def _ask_overwrite_all(_paths: list[str]) -> bool:
        raise AssertionError("Bulk confirmation must not be called in per-file mode.")

    def _ask_output_path(suggested_output: str) -> str | None:
        suggested_outputs.append(suggested_output)
        if suggested_output.endswith("/a.xlsx"):
            return "C:/out/custom_a.xlsx"
        return None

    resolved = resolve_overwrite_conflicts(
        conflicts,
        per_file_native_limit=2,
        ask_overwrite_all=_ask_overwrite_all,
        ask_output_path=_ask_output_path,
    )

    assert suggested_outputs == ["C:/out/a.xlsx", "C:/out/b.xlsx"]
    assert resolved.retry_sources == ["C:/src/a.xlsx"]
    assert resolved.output_path_overrides == {"C:/src/a.xlsx": "C:/out/custom_a.xlsx"}


def test_resolve_overwrite_conflicts_deduplicates_sources_by_canonical_key() -> None:
    conflicts = [
        OverwriteConflict(source_path="C:/SRC/A.xlsx", output_path="C:/out/a.xlsx"),
        OverwriteConflict(source_path="c:/src/a.xlsx", output_path="C:/out/other_a.xlsx"),
        OverwriteConflict(source_path="C:/src/b.xlsx", output_path="C:/out/b.xlsx"),
    ]

    def _ask_overwrite_all(_paths: list[str]) -> bool:
        return True

    def _ask_output_path(_suggested_output: str) -> str | None:
        return None

    resolved = resolve_overwrite_conflicts(
        conflicts,
        per_file_native_limit=1,
        ask_overwrite_all=_ask_overwrite_all,
        ask_output_path=_ask_output_path,
    )

    assert resolved.retry_sources == ["C:/SRC/A.xlsx", "C:/src/b.xlsx"]
    assert resolved.output_path_overrides == {
        "C:/SRC/A.xlsx": "C:/out/a.xlsx",
        "C:/src/b.xlsx": "C:/out/b.xlsx",
    }
