from __future__ import annotations

import ast
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
_MAX_TEMPLATE_ENGINE_LINES = 900


def _imports_for(path: Path) -> list[str]:
    """Imports for.
    
    Args:
        path: Parameter value (Path).
    
    Returns:
        list[str]: Return value.
    
    Raises:
        None.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imports.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imports.append(node.module)
    return imports


def test_services_layer_does_not_import_ui_modules() -> None:
    """Test services layer does not import ui modules.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    imports = _imports_for(service_file)
    if not (not any(name == "modules" or name.startswith("modules.") for name in imports)):
        raise AssertionError('assertion failed')


def test_template_version_module_does_not_import_engine_module() -> None:
    """Test template version module does not import engine module.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    version_file = REPO_ROOT / "domain" / "template_versions" / "course_setup_v2.py"
    imports = _imports_for(version_file)
    if not ("modules.instructor.instructor_template_engine" not in imports):
        raise AssertionError('assertion failed')


def test_domain_instructor_engine_does_not_import_ui_modules() -> None:
    """Test domain instructor engine does not import ui modules.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    if not (not (REPO_ROOT / "domain" / "instructor_engine.py").exists()):
        raise AssertionError('assertion failed')
    if not (not (REPO_ROOT / "domain" / "instructor_template_engine.py").exists()):
        raise AssertionError('assertion failed')
    if not (not (REPO_ROOT / "domain" / "instructor_report_engine.py").exists()):
        raise AssertionError('assertion failed')


def test_template_engine_stays_within_size_budget() -> None:
    """Test template engine stays within size budget.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    router_file = REPO_ROOT / "domain" / "template_strategy_router.py"
    line_count = len(router_file.read_text(encoding="utf-8").splitlines())
    if not (line_count <= _MAX_TEMPLATE_ENGINE_LINES):
        raise AssertionError('assertion failed')


def test_sheetops_module_does_not_import_ui_or_service_layers() -> None:
    """Test sheetops module does not import ui or service layers.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    sheetops_file = (
        REPO_ROOT / "domain" / "template_versions" / "course_setup_v2_impl" / "instructor_engine_sheetops.py"
    )
    imports = _imports_for(sheetops_file)
    if not (not any(name == "modules" or name.startswith("modules.") for name in imports)):
        raise AssertionError('assertion failed')
    if not (not any(name == "services" or name.startswith("services.") for name in imports)):
        raise AssertionError('assertion failed')


def test_service_layer_does_not_define_atomic_copy_helper() -> None:
    """Test service layer does not define atomic copy helper.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    content = service_file.read_text(encoding="utf-8")
    if not ("def _atomic_copy_file(" not in content):
        raise AssertionError('assertion failed')


def test_service_layer_does_not_import_removed_instructor_wrappers() -> None:
    """Test service layer does not import removed instructor wrappers.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    imports = _imports_for(service_file)
    if not ("domain.instructor_engine" not in imports):
        raise AssertionError('assertion failed')
    if not ("domain.instructor_template_engine" not in imports):
        raise AssertionError('assertion failed')
    if not ("domain.instructor_report_engine" not in imports):
        raise AssertionError('assertion failed')


def test_utils_does_not_own_workbook_integrity_rules() -> None:
    """Test utils does not own workbook integrity rules.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    utils_file = REPO_ROOT / "common" / "utils.py"
    content = utils_file.read_text(encoding="utf-8")
    if not ("def read_valid_template_id_from_system_hash_sheet(" not in content):
        raise AssertionError('assertion failed')
    if not ("def read_template_id_from_system_hash_sheet_if_valid(" not in content):
        raise AssertionError('assertion failed')
    if not ("def add_system_hash_sheet(" not in content):
        raise AssertionError('assertion failed')
    if not ("def add_system_layout_sheet(" not in content):
        raise AssertionError('assertion failed')
    if not ("def copy_system_hash_sheet(" not in content):
        raise AssertionError('assertion failed')


def test_template_strategy_router_uses_shared_workbook_integrity_package() -> None:
    """Test template strategy router uses shared workbook integrity package.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    router_file = REPO_ROOT / "domain" / "template_strategy_router.py"
    content = router_file.read_text(encoding="utf-8")
    if "from common.workbook_integrity import (" not in content:
        raise AssertionError('assertion failed')


def test_instructor_module_uses_shared_workbook_output_resolution_helper() -> None:
    """Test instructor module uses shared workbook output resolution helper.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    module_file = REPO_ROOT / "modules" / "instructor_module.py"
    content = module_file.read_text(encoding="utf-8")
    if "from common.workbook_output_resolution import (" not in content:
        raise AssertionError('assertion failed')
    if "extract_overwrite_conflicts_from_generation_result" not in content:
        raise AssertionError('assertion failed')
    if "resolve_overwrite_conflicts" not in content:
        raise AssertionError('assertion failed')
