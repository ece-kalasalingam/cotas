"""CO attainment CIP JSON payload builder for COURSE_SETUP_V2.

Produces a compact payload with short field codes.
The Cloudflare Worker holds a code dictionary that expands these codes
for Gemini — so the EXE sends minimum tokens.
"""

from __future__ import annotations

from common.constants import DIRECT_RATIO, INDIRECT_RATIO
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_TOTAL_STUDENTS_KEY,
)
from common.utils import normalize
from domain.template_versions.course_setup_v2_impl.co_description_template_validator import (
    CoDescriptionRecord,
)

_BLOOMS_CODE: dict[str, int] = {
    "REMEMBER": 1,
    "UNDERSTAND": 2,
    "APPLY": 3,
    "ANALYSE": 4,
    "ANALYZE": 4,
    "EVALUATE": 5,
    "CREATE": 6,
}


def _blooms(domain_level: str) -> int | str:
    """Return Bloom's numeric code (1-6) or the original string if unrecognised."""
    upper = domain_level.upper()
    for keyword, code in _BLOOMS_CODE.items():
        if keyword in upper:
            return code
    return domain_level


def _meta(metadata: dict[str, str], key: str) -> str:
    return str(metadata.get(normalize(key), "")).strip()


def _parse_pct(value: object) -> float | None:
    s = str(value).strip().rstrip("%")
    try:
        return float(s)
    except ValueError:
        return None


def _dist(attended: int, level_counts: dict[int, int]) -> list[int]:
    """Return [below_L1%, at_L1%, at_L2%, at_L3%] as integer percentages."""
    def _pct(count: int) -> int:
        return round(count / attended * 100) if attended > 0 else 0

    return [
        _pct(level_counts.get(0, 0)),
        _pct(level_counts.get(1, 0)),
        _pct(level_counts.get(2, 0)),
        _pct(level_counts.get(3, 0)),
    ]


def build_cip_payload(
    *,
    metadata: dict[str, str],
    thresholds: tuple[float, float, float],
    co_attainment_percent: float,
    co_attainment_level: int,
    total_outcomes: int,
    co_rows: list[dict[str, str]],
    co_level_data: dict[int, tuple[int, dict[int, int]]],
    assessments: list[dict[str, object]],
    co_description_records: list[CoDescriptionRecord],
) -> dict[str, object]:
    """Build the compact Gemini CIP payload from already-computed CO attainment data.

    co_level_data: maps CO index (1-based) to (attended, {level: student_count}).
    assessments: caller-built list of {name, wt, d} dicts.
    co_description_records: ordered by CO index 1..total_outcomes.
    """
    raw_students = _meta(metadata, COURSE_METADATA_TOTAL_STUDENTS_KEY)
    try:
        total_students = int(float(raw_students)) if raw_students else 0
    except ValueError:
        total_students = 0
    # Fall back to max attended count when metadata omits total_students.
    if total_students <= 0 and co_level_data:
        total_students = max((attended for attended, _ in co_level_data.values()), default=0)

    l1, l2, l3 = thresholds
    attained_count = sum(
        1 for row in co_rows if str(row.get("result", "")).strip() == "Attained"
    )

    cos: list[dict[str, object]] = []
    for i, row in enumerate(co_rows):
        co_index = i + 1
        attended, level_counts = co_level_data.get(co_index, (0, {}))
        desc = co_description_records[i] if i < len(co_description_records) else None
        result_raw = str(row.get("result", "")).strip()
        da_val = _parse_pct(row.get("direct")) or 0.0
        ia_val = _parse_pct(row.get("indirect")) or 0.0
        cos.append(
            {
                "id": str(row.get("co", f"CO{co_index}")).strip(),
                "desc": desc.description if desc else "",
                "bl": _blooms(desc.domain_level) if desc else "",
                "topics": desc.topics if desc else "",
                "da": da_val,
                "ia": ia_val,
                "avg": round(DIRECT_RATIO * da_val + INDIRECT_RATIO * ia_val, 2),
                "att_pct": _parse_pct(row.get("overall")) or 0.0,
                "st": "A" if result_raw == "Attained" else "NA",
                "sf": _parse_pct(row.get("shortfall")) or 0.0,
                "dist": _dist(attended, level_counts),
            }
        )

    return {
        "course": {
            "code": _meta(metadata, COURSE_METADATA_COURSE_CODE_KEY),
            "sem": _meta(metadata, COURSE_METADATA_SEMESTER_KEY),
            "ay": _meta(metadata, COURSE_METADATA_ACADEMIC_YEAR_KEY),
            "students": total_students,
        },
        "policy": {
            "tgt_pct": round(co_attainment_percent, 2),
            "tgt_lvl": f"L{co_attainment_level}",
            "thresh": [round(l1, 2), round(l2, 2), round(l3, 2)],
            "d_wt": round(DIRECT_RATIO * 100),
            "i_wt": round(INDIRECT_RATIO * 100),
        },
        "summary": {
            "total": total_outcomes,
            "att": attained_count,
            "not_att": total_outcomes - attained_count,
        },
        "assessments": assessments,
        "cos": cos,
    }


__all__ = ["build_cip_payload"]
