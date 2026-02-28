from typing import Any, Dict, List


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    raw = str(value).strip()
    if not raw:
        return None

    raw = raw.replace(",", "")
    if raw.endswith("%"):
        raw = raw[:-1].strip()

    try:
        return float(raw)
    except Exception:
        return None


def _normalize_direct_flag(value: Any) -> str:
    token = "" if value is None else str(value).strip().lower()
    if not token:
        return ""

    yes_tokens = {"yes", "y", "true", "1", "direct", "d"}
    no_tokens = {"no", "n", "false", "0", "indirect", "i"}

    if token in yes_tokens:
        return "YES"
    if token in no_tokens:
        return "NO"
    return ""


def check_conditional_weight_sum(engine: Any, p: Dict) -> List[str]:
    sheet_name = engine.bp.key_map.get(p["sheet_key"])
    weight_col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['weight_col_key']}")
    direct_col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['direct_col_key']}")

    w_idx = engine.get_col_idx(sheet_name, weight_col_name)
    d_idx = engine.get_col_idx(sheet_name, direct_col_name)

    if w_idx == -1 or d_idx == -1:
        return [f"Required columns not found in {sheet_name}"]

    rows = engine.data_store.get(sheet_name, [])
    actual_sum = 0.0
    target_condition = _normalize_direct_flag(p["condition_val"]) or str(p["condition_val"]).strip().upper()
    unknown_flags: set[str] = set()

    for row in rows:
        if d_idx >= len(row):
            continue

        normalized_flag = _normalize_direct_flag(row[d_idx])
        raw_flag = "" if row[d_idx] is None else str(row[d_idx]).strip()

        if not normalized_flag:
            if raw_flag:
                unknown_flags.add(raw_flag)
            continue

        if normalized_flag != target_condition:
            continue

        if w_idx >= len(row):
            continue

        parsed = _to_float(row[w_idx])
        if parsed is None:
            continue

        actual_sum += parsed

    if abs(actual_sum - float(p["target"])) > 0.01:
        msg = (
            f"{sheet_name}: total weight for Direct='{p['condition_val']}' "
            f"is {round(actual_sum, 2)}, expected {p['target']}"
        )
        if unknown_flags:
            samples = ", ".join(sorted(list(unknown_flags))[:5])
            msg += f". Unrecognized Direct values found: {samples}"
        return [msg]

    return []


def check_multiple_columns_empty(engine: Any, p: Dict) -> List[str]:
    sheet_name = engine.bp.key_map.get(p["sheet_key"])
    col_indices = []

    for col_key in p["col_keys"]:
        col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{col_key}")
        idx = engine.get_col_idx(sheet_name, col_name)
        if idx == -1:
            return [f"Column {col_name} not found in {sheet_name}"]
        col_indices.append(idx)

    rows = engine.data_store.get(sheet_name, [])
    for i, row in enumerate(rows, start=1):
        if any(idx >= len(row) or row[idx] is None or str(row[idx]).strip() == "" for idx in col_indices):
            return [f"{sheet_name}: row {i} has empty required column(s)"]

    return []


def check_cross_workbook_sync(engine: Any, p: Dict) -> List[str]:
    _ = engine
    _ = p
    return []
