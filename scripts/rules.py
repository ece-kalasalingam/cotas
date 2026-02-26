from dataclasses import dataclass
from typing import Callable, Dict, Any, List, Optional

@dataclass
class BusinessRule:
    rule_id: str
    scope: str  # "INTRA", "INTER", "CROSS"
    logic_fn: Callable
    # Maps Version ID -> Logic Parameters
    versioned_params: Dict[str, Dict[str, Any]]

    def run(self, engine: Any, external_engine: Optional[Any] = None) -> List[str]:
        # Version Gate
        params = self.versioned_params.get(engine.bp.type_id)
        if not params:
            return []

        if external_engine:
            params['external_engine'] = external_engine

        return self.logic_fn(engine, params)