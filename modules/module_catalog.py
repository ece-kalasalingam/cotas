"""Default module plugin catalog for the main activity bar."""

from __future__ import annotations

from common.module_plugins import ModulePluginSpec, lazy_module_class


def build_module_catalog() -> tuple[ModulePluginSpec, ...]:
    return (
        ModulePluginSpec(
            key="instructor",
            title_key="module.instructor",
            icon_path="assets/co_section.svg",
            class_loader=lazy_module_class("modules.instructor_module", "InstructorModule"),
        ),
        ModulePluginSpec(
            key="coordinator",
            title_key="module.coordinator_short",
            icon_path="assets/co_course.svg",
            class_loader=lazy_module_class("modules.coordinator_module", "CoordinatorModule"),
        ),
        ModulePluginSpec(
            key="po_analysis",
            title_key="module.po_analysis",
            icon_path="assets/po.svg",
            class_loader=lazy_module_class("modules.po_analysis_module", "POAnalysisModule"),
        ),
        ModulePluginSpec(
            key="co_analysis",
            title_key="module.co_analysis",
            icon_path="assets/co_course.svg",
            class_loader=lazy_module_class("modules.co_analysis_module", "COAnalysisModule"),
        ),
        ModulePluginSpec(
            key="help",
            title_key="nav.help",
            icon_path="assets/help.svg",
            class_loader=lazy_module_class("modules.help_module", "HelpModule"),
        ),
        ModulePluginSpec(
            key="about",
            title_key="nav.about",
            icon_path="assets/about.svg",
            class_loader=lazy_module_class("modules.about_module", "AboutModule"),
        ),
    )

