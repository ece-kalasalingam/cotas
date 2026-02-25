import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Callable
from scripts.sheet_schema import WorkbookBlueprint, ValidationRule
from scripts.utils import normalize
from scripts.exceptions import ValidationError

class UniversalEngine:
    """
    ONE CLASS FOR ALL:
    1. Extraction Unit: Converts Excel files into native Python List-of-Lists.
    2. Structural Unit: Performs O(1) cell validation using pre-indexed hash-maps.
    3. Logic Unit: Triggers the Ease-Chain for workbook-specific business rules.
    """
    def __init__(self, blueprint: WorkbookBlueprint):
        self.bp = blueprint
        self.errors: List[str] = []
        # Native Python data store (Fastest access for O(1) logic)
        self.data_store: Dict[str, List[List[Any]]] = {}

    def load_from_file(self, filepath: str) -> bool:
        """Technical Unit: Reads Excel into memory. Separated for modularity."""
        self.errors = []
        try:
            with pd.ExcelFile(filepath) as xls:
                for sheet_schema in self.bp.sheets:
                    if sheet_schema.name not in xls.sheet_names:
                        self.errors.append(f"Missing required sheet: {sheet_schema.name}")
                        continue
                    
                    # Convert to List-of-Lists (O(N) extraction)
                    # We replace NaN with None to keep the O(1) logic unit clean
                    df = pd.read_excel(xls, sheet_name=sheet_schema.name, header=None)
                    self.data_store[sheet_schema.name] = df.replace({np.nan: None}).values.tolist()
            
            return len(self.errors) == 0
        except Exception as e:
            self.errors.append(f"File Access Error: {str(e)}")
            return False

    def load_from_mock(self, mock_data: Dict[str, List[List[Any]]]):
        """Testing Unit: Allows O(1) validation testing without a physical file."""
        self.data_store = mock_data

    def run_validation(self, logic_chain: Optional[Callable[[Dict], List[str]]] = None) -> bool:
        """
        The Ease-Chain Entry Point:
        1. Validates structural rules defined in the blueprint (O(1)).
        2. Validates business logic rules defined in the logic_chain.
        """
        # 1. Structural Pass (Cell level)
        if not self._validate_structure():
            return False

        # 2. Business Pass (Cross-sheet / Logic level)
        if logic_chain:
            business_errors = logic_chain(self.data_store)
            if business_errors:
                self.errors.extend(business_errors)
                return False

        return True

    def _validate_structure(self) -> bool:
        """
        Core O(1) logic. Pre-indexes rules so that checking a cell 
        is a constant-time dictionary lookup.
        """
        for sheet_schema in self.bp.sheets:
            rows = self.data_store.get(sheet_schema.name, [])
            if not rows: continue

            # --- PRE-COMPILE RULE MAP (O1 Step) ---
            # Dictionary: Column_Index -> List[ValidationRule]
            col_rule_map: Dict[int, List[ValidationRule]] = {}
            for rule in sheet_schema.validations:
                # Cache list sources as sets for O(1) 'in' checks
                if 'source' in rule.options and isinstance(rule.options['source'], list):
                    rule.options['_set_cache'] = {normalize(x) for x in rule.options['source']}
                
                for col in range(rule.first_col, rule.last_col + 1):
                    col_rule_map.setdefault(col, []).append(rule)

            # --- VALIDATION LOOP ---
            # Height depends on the blueprint header matrix
            header_height = len(sheet_schema.header_matrix)
            
            for r_idx, row in enumerate(rows[header_height:], start=header_height + 1):
                for col_idx, rules in col_rule_map.items():
                    if col_idx >= len(row): continue
                    
                    cell_val = row[col_idx]
                    norm_val = normalize(cell_val)
                    
                    for rule in rules:
                        if not self._execute_rule(norm_val, rule):
                            self.errors.append(
                                f"[{sheet_schema.name}] Row {r_idx}, Col {col_idx+1}: "
                                f"'{cell_val}' fails {rule.options.get('validate')} validation."
                            )
        
        return len(self.errors) == 0

    def _execute_rule(self, norm_val: str, rule: ValidationRule) -> bool:
        """Actual logic check. Every branch here is O(1)."""
        opts = rule.options
        v_type = opts.get('validate')

        if v_type == 'list':
            return norm_val in opts.get('_set_cache', set())
        
        if v_type in ('decimal', 'whole'):
            if not norm_val: return True # Handle empty based on blueprint
            try:
                num = float(norm_val)
                criteria = opts.get('criteria')
                limit = float(opts.get('value', 0))
                
                if criteria == 'greater than': return num > limit
                if criteria == 'between':
                    return float(opts.get('min', 0)) <= num <= float(opts.get('max', 100))
            except (ValueError, TypeError):
                return False
                
        return True

    def get_data_store(self) -> Dict[str, List[List[Any]]]:
        """Provides the Ease-Chain with raw data for business logic."""
        return self.data_store
    def get_sheet_as_dict(self, sheet_name: str) -> Dict[str, Any]:
        """Converts a two-column sheet (Key, Value) into a dictionary."""
        rows = self.data_store.get(sheet_name, [])
        # We normalize keys so 'Course Code' becomes 'course_code'
        return {normalize(str(r[0])): r[1] for r in rows if len(r) >= 2}