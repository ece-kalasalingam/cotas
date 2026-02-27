# scripts/rules.py
from dataclasses import dataclass
from typing import Callable, Dict, Any, List, Optional

@dataclass
class BusinessRule:
    rule_id: str
    scope: str  # "INTRA", "INTER", "CROSS"
    logic_fn: Callable
    versioned_params: Dict[str, Dict[str, Any]]

    def run(self, engine: Any, external_engine: Optional[Any] = None) -> List[str]:
        # Get parameters specific to the template version currently loaded in the engine
        given_params = self.versioned_params.get(engine.bp.type_id)
        if not given_params:
            return []
        
        params = given_params.copy()
        if external_engine:
            params['external_engine'] = external_engine

        # We pass 'engine' so the logic_fn can use engine.get_col_idx() or engine.bp
        return self.logic_fn(engine, params)