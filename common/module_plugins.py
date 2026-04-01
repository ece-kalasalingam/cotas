"""Plugin contracts for MainWindow activity modules."""

from __future__ import annotations

import importlib
from collections.abc import Callable
from dataclasses import dataclass

from PySide6.QtWidgets import QWidget

ModuleClass = type[QWidget]
ModuleClassLoader = Callable[[], ModuleClass]


@dataclass(frozen=True, slots=True)
class ModulePluginSpec:
    key: str
    title_key: str
    icon_path: str
    class_loader: ModuleClassLoader
    show_in_activity_bar: bool = True


def lazy_module_class(module_path: str, class_name: str) -> ModuleClassLoader:
    """Lazy module class.
    
    Args:
        module_path: Parameter value (str).
        class_name: Parameter value (str).
    
    Returns:
        ModuleClassLoader: Return value.
    
    Raises:
        None.
    """
    def _load() -> ModuleClass:
        """Load.
        
        Args:
            None.
        
        Returns:
            ModuleClass: Return value.
        
        Raises:
            None.
        """
        module = importlib.import_module(module_path)
        loaded = getattr(module, class_name)
        return loaded

    return _load

