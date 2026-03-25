from __future__ import annotations

import ast
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
_MAX_TEMPLATE_ENGINE_LINES = 900
_MAX_COORDINATOR_MODULE_LINES = 1425
_MAX_INSTRUCTOR_REPORT_ENGINE_LINES = 1075


def _imports_for(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imports.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imports.append(node.module)
    return imports


def test_services_layer_does_not_import_ui_modules() -> None:
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    imports = _imports_for(service_file)
    assert not any(name == "modules" or name.startswith("modules.") for name in imports)


def test_template_version_module_does_not_import_engine_module() -> None:
    version_file = REPO_ROOT / "domain" / "template_versions" / "course_setup_v1.py"
    imports = _imports_for(version_file)
    assert "modules.instructor.instructor_template_engine" not in imports


def test_domain_instructor_engine_does_not_import_ui_modules() -> None:
    engine_file = REPO_ROOT / "domain" / "instructor_template_engine.py"
    imports = _imports_for(engine_file)
    assert not any(name == "modules" or name.startswith("modules.") for name in imports)


def test_domain_coordinator_engine_does_not_import_ui_modules() -> None:
    engine_file = REPO_ROOT / "domain" / "coordinator_engine.py"
    imports = _imports_for(engine_file)
    assert not any(name == "modules" or name.startswith("modules.") for name in imports)


def test_template_engine_stays_within_size_budget() -> None:
    engine_file = REPO_ROOT / "domain" / "instructor_template_engine.py"
    line_count = len(engine_file.read_text(encoding="utf-8").splitlines())
    assert line_count <= _MAX_TEMPLATE_ENGINE_LINES


def test_coordinator_module_stays_within_size_budget() -> None:
    coordinator_file = REPO_ROOT / "modules" / "coordinator_module.py"
    line_count = len(coordinator_file.read_text(encoding="utf-8").splitlines())
    assert line_count <= _MAX_COORDINATOR_MODULE_LINES


def test_instructor_report_engine_stays_within_size_budget() -> None:
    report_engine_file = REPO_ROOT / "domain" / "instructor_report_engine.py"
    line_count = len(report_engine_file.read_text(encoding="utf-8").splitlines())
    assert line_count <= _MAX_INSTRUCTOR_REPORT_ENGINE_LINES


def test_sheetops_module_does_not_import_ui_or_service_layers() -> None:
    sheetops_file = (
        REPO_ROOT / "domain" / "template_versions" / "course_setup_v2_impl" / "instructor_engine_sheetops.py"
    )
    imports = _imports_for(sheetops_file)
    assert not any(name == "modules" or name.startswith("modules.") for name in imports)
    assert not any(name == "services" or name.startswith("services.") for name in imports)


def test_service_layer_does_not_define_atomic_copy_helper() -> None:
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    content = service_file.read_text(encoding="utf-8")
    assert "def _atomic_copy_file(" not in content


def test_utils_does_not_own_workbook_integrity_rules() -> None:
    utils_file = REPO_ROOT / "common" / "utils.py"
    content = utils_file.read_text(encoding="utf-8")
    assert "def read_valid_template_id_from_system_hash_sheet(" not in content
    assert "def read_template_id_from_system_hash_sheet_if_valid(" not in content
    assert "def add_system_hash_sheet(" not in content
    assert "def add_system_layout_sheet(" not in content
    assert "def copy_system_hash_sheet(" not in content


def test_template_strategy_router_uses_shared_workbook_integrity_package() -> None:
    router_file = REPO_ROOT / "domain" / "template_strategy_router.py"
    content = router_file.read_text(encoding="utf-8")
    assert "from common.workbook_integrity import (" in content
