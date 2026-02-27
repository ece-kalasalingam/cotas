# scripts/logic_library.py
from typing import Any, Dict, List

def check_conditional_weight_sum(engine: Any, p: Dict) -> List[str]:
    """
    Checks if the sum of 'Weight (%)' matches a target based on the 'Direct' column.
    p = {
        'sheet_key': 'Sheet1',
        'weight_col_key': 'Weight (%)',
        'direct_col_key': 'Direct',
        'condition_val': 'YES',
        'target': 100
    }
    """
    # 1. Resolve Sheet and Column Names from Blueprint mapping
    sheet_name = engine.bp.key_map.get(p['sheet_key'])
    weight_col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['weight_col_key']}")
    direct_col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['direct_col_key']}")

    # 2. Get Column Indices
    w_idx = engine.get_col_idx(sheet_name, weight_col_name)
    d_idx = engine.get_col_idx(sheet_name, direct_col_name)

    if w_idx == -1 or d_idx == -1:
        return [f"Required columns not found in {sheet_name}"]

    # 3. Iterate through data_store (Row Iteration)
    rows = engine.data_store.get(sheet_name, [])
    actual_sum = 0.0
    target_condition = str(p['condition_val']).strip().upper()

    for row in rows:
        try:
            # Check the 'Direct' column value
            if str(row[d_idx]).strip().upper() == target_condition:
                # Add the 'Weight (%)' value
                val = row[w_idx]
                if val is not None:
                    actual_sum += float(val)
        except (ValueError, TypeError):
            continue # Skip non-numeric weights

    # 4. Validate against target (with float tolerance)
    if abs(actual_sum - p['target']) > 0.01:
        return [
            f"{sheet_name}: The total weight for Direct='{p['condition_val']}' "
            f"is {round(actual_sum, 2)}, but expected {p['target']}."
        ]

    return []

def check_column_sum(engine: Any, p: Dict) -> List[str]:
    """Logic-only: Checks if a column total matches a target."""
    sheet_name = engine.bp.key_map.get(p['sheet_key'])
    col_name = engine.bp.key_map.get(f"{p['sheet_key']}.{p['col_key']}")
    idx = engine.get_col_idx(sheet_name, col_name)

    if idx == -1: return [f"Column {col_name} not found."]

    # Use the lightweight engine's pre-calculated cache
    actual_sum = engine._col_cache.get(sheet_name, {}).get(f"SUM_{idx}", 0.0)
    
    if abs(actual_sum - p['target']) > 0.01:
        return [f"{sheet_name}: {p['col_key']} sum is {actual_sum}, expected {p['target']}"]
    return []
def check_cross_workbook_sync(engine: Any, p: Dict) -> List[str]:
    """Syncs current file against an external one (e.g., Student IDs)."""
    ext_engine = p.get('external_engine')
    if not ext_engine: 
        return ["Sync Error: Master reference file missing."]

    # 1. Resolve Current Workbook Details
    c_sheet = engine.bp.key_map.get(p['curr_sheet'])
    c_col = engine.bp.key_map.get(f"{p['curr_sheet']}.{p['curr_col']}")
    c_idx = engine.get_col_idx(c_sheet, c_col)

    # 2. Resolve External Workbook Details
    e_sheet = ext_engine.bp.key_map.get(p['ext_sheet'])
    e_col = ext_engine.bp.key_map.get(f"{p['ext_sheet']}.{p['ext_col']}")
    e_idx = ext_engine.get_col_idx(e_sheet, e_col)

    # Safety Check: Ensure columns were found
    if c_idx == -1 or e_idx == -1:
        return [f"Sync Error: Mapping failed for {c_col} or {e_col}."]

    # 3. Extract IDs with cleaning (Handles None and leading/trailing spaces)
    def get_clean_ids(target_engine, sheet, col_idx):
        data = target_engine.data_store.get(sheet, [])
        ids = set()
        for row in data:
            if col_idx < len(row) and row[col_idx] is not None:
                val = str(row[col_idx]).strip()
                if val: ids.add(val)
        return ids

    curr_ids = get_clean_ids(engine, c_sheet, c_idx)
    ext_ids = get_clean_ids(ext_engine, e_sheet, e_idx)

    # 4. Find Mismatches
    diff = curr_ids - ext_ids
    if diff:
        return [f"ID Mismatch: {len(diff)} entries in current file (e.g., {list(diff)[:3]}) not found in Master."]
    
    return []