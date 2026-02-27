# scripts/rules.py
from dataclasses import dataclass
from typing import Callable, Dict, Any, List, Optional

ALLOWED_SCOPES = {"INTRA", "INTER", "CROSS"}


def _canon(value: Any) -> str:
    return "" if value is None else str(value).strip().lower()


@dataclass
class BusinessRule:
    rule_id: str
    scope: str  # "INTRA", "INTER", "CROSS"
    logic_fn: Callable
    versioned_params: Dict[str, Dict[str, Any]]

    def run(self, engine: Any, external_engine: Optional[Any] = None) -> List[str]:
        scope = str(self.scope).strip().upper()
        if scope not in ALLOWED_SCOPES:
            return [f"Rule {self.rule_id}: unsupported scope '{self.scope}'"]

        # Get parameters specific to the template version currently loaded in the engine
        given_params = self.versioned_params.get(engine.bp.type_id)
        if not given_params:
            return []

        params = given_params.copy()
        if scope == "CROSS":
            if external_engine is None:
                return [f"Rule {self.rule_id}: external workbook required for CROSS scope"]

            target_type_id = params.get("target_type_id")
            if not target_type_id:
                return [f"Rule {self.rule_id}: target_type_id is required for CROSS scope"]

            ext_bp = getattr(external_engine, "bp", None)
            ext_type_id = getattr(ext_bp, "type_id", None) if ext_bp else None
            if _canon(ext_type_id) != _canon(target_type_id):
                return [
                    f"Rule {self.rule_id}: external workbook type mismatch "
                    f"(expected '{target_type_id}', got '{ext_type_id}')"
                ]

            params['external_engine'] = external_engine

        # We pass 'engine' so the logic_fn can use engine.get_col_idx() or engine.bp
        errors = self.logic_fn(engine, params) or []
        return [f"Rule {self.rule_id}: {err}" for err in errors]
