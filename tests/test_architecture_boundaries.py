from __future__ import annotations

import ast
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
_MAX_TEMPLATE_ENGINE_LINES = 900


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


def test_template_engine_stays_within_size_budget() -> None:
    engine_file = REPO_ROOT / "domain" / "instructor_template_engine.py"
    line_count = len(engine_file.read_text(encoding="utf-8").splitlines())
    assert line_count <= _MAX_TEMPLATE_ENGINE_LINES


def test_sheetops_module_does_not_import_ui_or_service_layers() -> None:
    sheetops_file = REPO_ROOT / "domain" / "instructor_template_engine_sheetops.py"
    imports = _imports_for(sheetops_file)
    assert not any(name == "modules" or name.startswith("modules.") for name in imports)
    assert not any(name == "services" or name.startswith("services.") for name in imports)


def test_service_layer_does_not_define_atomic_copy_helper() -> None:
    service_file = REPO_ROOT / "services" / "instructor_workflow_service.py"
    content = service_file.read_text(encoding="utf-8")
    assert "def _atomic_copy_file(" not in content
