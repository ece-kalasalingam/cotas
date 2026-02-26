from typing import List, Any
from scripts.utils import normalize
from scripts.exceptions import ValidationError, SystemError

def validate_course_setup_logic(engine: Any) -> List[str]:
    """
    The Master Entry Point. 
    Returns a list of strings for user-facing errors.
    Raises SystemError for blueprint/logic bugs.
    """
    errors = []
    
    # 1. Check Structure first (O1 Cell Rules)
    errors.extend(run_structural_checks(engine))
    
    # 2. Check Business Logic (only if structure is valid)
    if not errors:
        # We wrap this in a try-block to catch critical logic stops
        try:
            errors.extend(check_weightage_sum(engine))
            errors.extend(check_co_range_validity(engine))
        except ValidationError as ve:
            errors.append(f"Fatal Data Error: {str(ve)}")
        except Exception as e:
            # If the code crashes, it's a SystemError (Blueprint/Logic bug)
            raise SystemError(f"Logic Validator crashed unexpectedly: {str(e)}")
            
    return errors

def run_structural_checks(engine: Any) -> List[str]:
    """Iterates through every cell and applies Blueprint rules in O(1)."""
    errors = []
    for sheet_schema in engine.bp.sheets:
        rows = engine.data_store.get(sheet_schema.name, [])
        if not rows:
            continue

        # Build Rule Map (The O1 Optimization)
        col_rule_map = {}
        for rule in sheet_schema.validations:
            if 'source' in rule.options:
                rule.options['_set_cache'] = {normalize(x) for x in rule.options['source']}
            for col in range(rule.first_col, rule.last_col + 1):
                col_rule_map.setdefault(col, []).append(rule)

        # Header height determines where data starts
        h_height = len(sheet_schema.header_matrix)
        for r_idx, row in enumerate(rows, start=h_height + 1):
            for col_idx, rules in col_rule_map.items():
                if col_idx >= len(row): continue
                
                val = row[col_idx]
                norm_val = normalize(val)
                
                for rule in rules:
                    if not execute_rule(norm_val, rule):
                        errors.append(f"[{sheet_schema.name}] Row {r_idx}, Col {col_idx+1}: '{val}' is invalid.")
    return errors

def execute_rule(norm_val: str, rule: Any) -> bool:
    """Atomic decision logic for a single cell."""
    opts = rule.options
    v_type = opts.get('validate')

    if v_type == 'list':
        return norm_val in opts.get('_set_cache', set())
    
    if v_type in ('decimal', 'whole'):
        if not norm_val or norm_val == "none": return True
        try:
            num = float(norm_val)
            criteria = opts.get('criteria')
            if criteria == 'greater than': return num > float(opts.get('value', 0))
            if criteria == 'between':
                return float(opts.get('min', 0)) <= num <= float(opts.get('max', 100))
        except: return False
    return True

# --- BUSINESS LOGIC UNITS ---

def check_weightage_sum(engine: Any) -> List[str]:
    sheet = "Assessment_Config"
    col_name = "Weight (%)"
    
    # Use our O(1) Cache lookup
    idx = engine.get_col_idx(sheet, col_name)
    
    # SystemError: The developer named the column wrong in the code or Blueprint
    if idx == -1:
        raise SystemError(f"Column '{col_name}' not found in cache for sheet '{sheet}'. Check Blueprint.")

    total = 0.0
    for row in engine.data_store.get(sheet, []):
        try:
            val = row[idx]
            if val: total += float(val)
        except (ValueError, TypeError):
            continue # Structural check already caught this

    if abs(total - 100.0) > 0.01:
        return [f"[{sheet}]: Total weightage must be 100% (Found: {total}%)."]
    
    return []

def check_co_range_validity(engine: Any) -> List[str]:
    """Logic for Cross-Sheet validation (e.g. COs vs Metadata)"""
    # implementation here...
    return []