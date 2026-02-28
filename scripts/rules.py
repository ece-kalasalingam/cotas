from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional


ALLOWED_SCOPES = {"INTRA", "INTER", "CROSS"}


def _canon(value: Any) -> str:
    return "" if value is None else str(value).strip().lower()


@dataclass
class BusinessRule:
    rule_id: str
    scope: str
    logic_fn: Callable
    versioned_params: Dict[str, Dict[str, Any]]

    def run(self, engine: Any, external_engine: Optional[Any] = None) -> List[str]:
        scope = str(self.scope).strip().upper()
        if scope not in ALLOWED_SCOPES:
            return [f"unsupported scope '{self.scope}'"]

        params = self.versioned_params.get(getattr(engine.bp, "type_id", None), {}).copy()
        if not params:
            return []

        if scope == "CROSS":
            if external_engine is None:
                return ["external workbook required"]
            target_type_id = params.get("target_type_id")
            ext_type_id = getattr(getattr(external_engine, "bp", None), "type_id", None)
            if target_type_id and _canon(target_type_id) != _canon(ext_type_id):
                return [f"external workbook type mismatch (expected {target_type_id}, got {ext_type_id})"]
            params["external_engine"] = external_engine

        return self.logic_fn(engine, params) or []
