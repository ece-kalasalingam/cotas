from typing import Any, Dict, List

def check_column_sum(engine: Any, p: Dict) -> List[str]:
    """Sums a column and checks against a target (e.g., Weightages = 100)."""
    # 1. Resolve physical names from the blueprint
    sheet_name = engine.bp.key_map.get(p['sheet_key'])
    col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['col_key']}")
    
    idx = engine.get_col_idx(sheet_name, col_name)
    if idx == -1: return [f"Error: {col_name} missing."]

    total = 0.0
    for row in engine.data_store.get(sheet_name, []):
        try:
            if row[idx] is not None: total += float(row[idx])
        except: continue

    if abs(total - p['target']) > 0.01:
        return [f"{sheet_name}: Total {p['col_key']} must be {p['target']}, found {total}"]
    return []

def check_cross_workbook_sync(engine: Any, p: Dict) -> List[str]:
    """Syncs current file against an external one (e.g., Student IDs)."""
    ext_engine = p.get('external_engine')
    if not ext_engine: return ["Sync Error: Master reference file missing."]

    # Current Workbook Resolve
    c_sheet = engine.bp.key_map[p['curr_sheet']]
    c_col = engine.bp.key_map[f"{p['curr_sheet']}.{p['curr_col']}"]
    c_idx = engine.get_col_idx(c_sheet, c_col)

    # External Workbook Resolve
    e_sheet = ext_engine.bp.key_map[p['ext_sheet']]
    e_col = ext_engine.bp.key_map[f"{p['ext_sheet']}.{p['ext_col']}"]
    e_idx = ext_engine.get_col_idx(e_sheet, e_col)

    curr_ids = {str(r[c_idx]) for r in engine.data_store[c_sheet] if r[c_idx]}
    ext_ids = {str(r[e_idx]) for r in ext_engine.data_store[e_sheet] if r[e_idx]}

    diff = curr_ids - ext_ids
    return [f"Unknown IDs in upload: {diff}"] if diff else []