"""Instructor CO module UI (single workflow: course template + marks template)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import cast

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from common.async_operation_runner import AsyncOperationRunner
from common.constants import (
    APP_NAME,
    ID_COURSE_SETUP,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
)
from common.error_catalog import validation_error_from_key
from common.drag_drop_file_widget import ManagedDropFileWidget
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.module_messages import build_status_message as _build_status_message
from common.module_messages import default_messages_namespace as _default_messages_namespace
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_runtime import ModuleRuntime
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputItem, OutputPanelData
from common.qt_jobs import run_in_background
from common.workbook_output_resolution import (
    extract_overwrite_conflicts_from_generation_result,
    resolve_overwrite_conflicts,
)
from common.i18n import t
from common.utils import canonical_path_key, log_process_message, resolve_dialog_start_path
from domain import BusyWorkflowState
from domain.template_strategy_router import (
    generate_workbook,
    generate_workbooks,
    validate_workbooks,
)

_logger = logging.getLogger(__name__)
_DOWNLOAD_COURSE_TEMPLATE_HREF = "download-course-template"


class _LogSink:
    def appendPlainText(self, _text: str) -> None:  # noqa: N802 - Qt-style name
        return

    def clear(self) -> None:
        return


class InstructorModule(QWidget):
    """Single-flow Instructor UI: download course template, upload course details, generate marks template."""

    status_changed = Signal(str)

    OUTPUT_LINKS = (
        ("instructor.links.course_details_generated", "course_template_path"),
        ("instructor.links.marks_template_generated", "marks_template_path"),
    )

    def __init__(self):
        super().__init__()

        self.course_template_path: str | None = None
        self.course_details_paths: list[str] = []
        self.marks_template_path: str | None = None
        self.marks_template_paths: list[str] = []
        self._uploaded_course_details_paths: list[str] = []
        self._validated_course_details_paths: list[str] = []
        self._ready_for_marks_generation = False
        self._syncing_drop_widget_files = False

        self.state = BusyWorkflowState()
        self._logger = _logger
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._is_closing = False
        self._ui_log_handler: logging.Handler | None = None
        self._user_log_entries: list[dict[str, object]] = []

        self._async_runner = AsyncOperationRunner(
            self,
            run_async=run_in_background,
            refresh_ui=lambda: self._refresh_ui(),
            should_refresh_ui=lambda: not self._is_closing,
        )
        self._runtime = ModuleRuntime(
            module=self,
            app_name=APP_NAME,
            logger=self._logger,
            async_runner=self._async_runner,
            messages_namespace_factory=_messages_namespace,
        )

        self._build_ui()
        self._setup_ui_logging()
        self._refresh_ui()

    def _build_ui(self) -> None:
        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="instructorTopRegion",
                footer_height=INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
            ),
        )
        top_pane = QWidget()
        top_layout = QHBoxLayout(top_pane)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        self._ui_engine.set_top_widget(top_pane)

        right_container = QWidget()
        right_container.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        right_scroll = QScrollArea()
        right_scroll.setObjectName("instructorRightScroll")
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.viewport().setObjectName("instructorRightScrollViewport")
        right_scroll.setWidget(right_container)
        top_layout.addWidget(right_scroll, 1)

        self.workflow_title = QLabel(t("instructor.workflow_title"))
        self.workflow_title.setObjectName("instructorRailTitle")
        right_layout.addWidget(self.workflow_title, 0)

        self.download_course_template_link = QLabel()
        self.download_course_template_link.setTextFormat(Qt.TextFormat.RichText)
        self.download_course_template_link.setTextInteractionFlags(
            Qt.TextInteractionFlag.LinksAccessibleByMouse | Qt.TextInteractionFlag.LinksAccessibleByKeyboard
        )
        self.download_course_template_link.setCursor(Qt.CursorShape.PointingHandCursor)
        self.download_course_template_link.setOpenExternalLinks(False)
        self.download_course_template_link.linkActivated.connect(self._on_download_course_template_link_activated)
        right_layout.addWidget(self.download_course_template_link, 0)

        self.course_details_drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("outputs.open_file"),
            open_folder_tooltip=t("outputs.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
        )
        self.course_details_drop_widget.files_dropped.connect(self._on_course_details_files_dropped)
        self.course_details_drop_widget.files_rejected.connect(self._on_course_details_files_rejected)
        self.course_details_drop_widget.browse_requested.connect(self._on_course_details_browse_requested)
        self.course_details_drop_widget.browse_requested.connect(self._upload_course_details_async)
        self.course_details_drop_widget.files_changed.connect(self._on_course_details_files_changed)
        self.course_details_drop_widget.clear_button.clicked.connect(self._on_clear_course_details_clicked)
        self.course_details_drop_widget.submit_requested.connect(self._prepare_marks_template_async)
        self.course_details_drop_widget.set_summary_text_builder(
            lambda count: t("instructor.drop.summary", count=count)
        )
        self.course_details_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.course_details_drop_widget.submit_button.setObjectName("primaryAction")
        self.course_details_drop_widget.submit_button.setCursor(Qt.CursorShape.PointingHandCursor)
        right_layout.addWidget(self.course_details_drop_widget, 1)

        self.generate_marks_template_action = self.course_details_drop_widget.submit_button

        self.user_log_view = _LogSink()
        self._ui_engine.set_footer_visible(False)

        self.shortcut_open_workbook = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_open_workbook.activated.connect(self._on_open_shortcut_activated)
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

    def _refresh_ui(self) -> None:
        self.workflow_title.setText(t("instructor.workflow_title"))
        self._set_download_course_template_link_enabled(not self.state.busy)
        self.course_details_drop_widget.set_submit_button_text(t("instructor.action.generate_marks_template"))
        self.course_details_drop_widget.set_clear_button_text(t("coordinator.clear_all"))
        self.course_details_drop_widget.set_summary_text_builder(
            lambda count: t("instructor.drop.summary", count=count)
        )
        self.course_details_drop_widget.set_submit_allowed(self._ready_for_marks_generation)

        enabled = not self.state.busy
        self.course_details_drop_widget.setEnabled(enabled)
        self.generate_marks_template_action.setEnabled(
            enabled and self._ready_for_marks_generation and bool(self.course_details_paths)
        )

    def retranslate_ui(self) -> None:
        self.course_details_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self._refresh_ui()

    def _output_items(self) -> tuple[OutputItem, ...]:
        rows: list[OutputItem] = []
        for label_key, attr in self.OUTPUT_LINKS:
            if attr == "marks_template_path":
                batch_paths = [path for path in self.marks_template_paths if path]
                if batch_paths:
                    rows.extend(OutputItem(label_key=label_key, path=path) for path in batch_paths)
                    continue
            value = getattr(self, attr)
            if isinstance(value, str) and value:
                rows.append(OutputItem(label_key=label_key, path=value))
        return tuple(rows)

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=self._output_items())

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self._ui_engine.set_footer_visible(False)

    def _set_download_course_template_link_enabled(self, enabled: bool) -> None:
        text = t("instructor.action.download_course_template")
        if enabled:
            self.download_course_template_link.setText(
                t(
                    "instructor.action.download_course_template_link_html",
                    href=_DOWNLOAD_COURSE_TEMPLATE_HREF,
                    text=text,
                )
            )
        else:
            self.download_course_template_link.setText(text)
        self.download_course_template_link.setEnabled(enabled)

    def _on_download_course_template_link_activated(self, _href: str) -> None:
        if self.state.busy:
            return
        self._download_course_template_async()

    def _on_open_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        self._upload_course_details_async()

    def _on_save_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        self._prepare_marks_template_async()

    def _on_course_details_browse_requested(self) -> None:
        self._publish_status_key("instructor.status.course_details_drop_browse_requested")

    def _on_course_details_files_dropped(self, dropped_files: list[str]) -> None:
        dropped_count = len([path for path in dropped_files if path])
        self._publish_status_key("instructor.status.course_details_drop_files_dropped", count=dropped_count)
        if self.state.busy:
            return
        selected_paths = [path for path in dropped_files if path]
        if not selected_paths:
            return
        self._upload_course_details_from_paths_async(selected_paths)

    def _on_course_details_files_rejected(self, files: list[str]) -> None:
        rejected_count = len([path for path in files if path])
        if rejected_count <= 0:
            return
        self._publish_status_key("instructor.status.course_details_drop_files_rejected", count=rejected_count)

    def _on_clear_course_details_clicked(self) -> None:
        if self.state.busy:
            return
        self.course_details_drop_widget.clear_files()

    def _on_course_details_files_changed(self, files: list[str]) -> None:
        if self.state.busy:
            return
        if self._syncing_drop_widget_files:
            return
        if files:
            current_keys = {canonical_path_key(path) for path in self._validated_course_details_paths}
            incoming_keys = {canonical_path_key(path) for path in files}
            if incoming_keys.issubset(current_keys):
                key_to_existing = {
                    canonical_path_key(path): path
                    for path in self._validated_course_details_paths
                }
                self.course_details_paths = [
                    key_to_existing[key]
                    for key in [canonical_path_key(path) for path in files]
                    if key in key_to_existing
                ]
                self._uploaded_course_details_paths = list(self.course_details_paths)
                self._validated_course_details_paths = list(self.course_details_paths)
                self._ready_for_marks_generation = bool(self.course_details_paths)
                self._refresh_ui()
                return

        had_state = bool(self._uploaded_course_details_paths or self._validated_course_details_paths)
        self._uploaded_course_details_paths = [path for path in files if path]
        self._validated_course_details_paths = []
        self.course_details_paths = []
        self.marks_template_paths = []
        self.marks_template_path = None
        self._ready_for_marks_generation = False
        if had_state:
            self._publish_status_key("instructor.status.course_details_replaced")
        self._refresh_ui()

    def _download_course_template_async(self) -> None:
        if self.state.busy:
            return

        start_dir = resolve_dialog_start_path(APP_NAME)
        default_path = str(Path(start_dir) / "Course_Details_Template.xlsx")
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            t("instructor.dialog.course_template.save_title"),
            default_path,
            t("instructor.dialog.filter.excel"),
        )
        if not output_path:
            return
        selected_output = Path(output_path)
        self._remember_dialog_dir_safe(str(selected_output))

        process_key = "instructor.log.process.generate_course_details_template"
        process_name = t(process_key)
        user_success_message, user_error_message = _localized_log_messages(process_key)
        token = CancellationToken()

        def _work() -> object:
            token.raise_if_cancelled()
            return generate_workbook(
                template_id=ID_COURSE_SETUP,
                output_path=selected_output,
                workbook_name=selected_output.name,
                workbook_kind="course_details_template",
                cancel_token=token,
            )

        def _on_success(result: object) -> None:
            output_value = str(
                getattr(result, "workbook_path", None)
                or getattr(result, "output_path", None)
                or selected_output
            ).strip()
            if not output_value:
                self._handle_async_failure(
                    validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="WORKBOOK_GENERATE_FAILED",
                        workbook_kind="course_details_template",
                    ),
                    process_name=process_name,
                    user_error_message=user_error_message,
                )
                return
            self.course_template_path = output_value
            self._publish_status_key("instructor.status.template_download_path_selected")
            log_process_message(
                process_name,
                logger=self._logger,
                success_message=f"{process_name} completed successfully.",
                user_success_message=user_success_message,
            )
            self._runtime.notify_message_key(
                "instructor.log.completed_process",
                channels=("toast",),
                kwargs={"process": t(process_key)},
                toast_title_key="instructor.msg.success_title",
                toast_level="success",
            )

        def _on_failure(exc: Exception) -> None:
            self._handle_async_failure(exc, process_name=process_name, user_error_message=user_error_message)

        self._start_async_operation(
            token=token,
            job_id=None,
            work=_work,
            on_success=_on_success,
            on_failure=_on_failure,
        )

    def _upload_course_details_async(self) -> None:
        if self.state.busy:
            return
        open_paths, _ = QFileDialog.getOpenFileNames(
            self,
            t("instructor.dialog.course_details.select_title"),
            resolve_dialog_start_path(APP_NAME),
            t("instructor.dialog.filter.excel_open"),
        )
        if not open_paths:
            return
        self._upload_course_details_from_paths_async(open_paths)

    def _upload_course_details_from_paths_async(self, open_paths: list[str]) -> None:
        if self.state.busy:
            return
        selected_paths = [path for path in open_paths if path]
        if not selected_paths:
            return
        had_previous_validated = bool(self._validated_course_details_paths)
        staged_uploaded_paths = self._merge_uploaded_paths(selected_paths)
        self._uploaded_course_details_paths = list(staged_uploaded_paths)
        self._validated_course_details_paths = []
        self.course_details_paths = []
        self.marks_template_paths = []
        self.marks_template_path = None
        self._ready_for_marks_generation = False
        self._set_course_details_widget_files(staged_uploaded_paths)
        self._refresh_ui()

        self._remember_dialog_dir_safe(selected_paths[0])
        process_key = "instructor.log.process.validate_course_details_workbook"
        process_name = t(process_key)
        user_success_message, user_error_message = _localized_log_messages(process_key)
        token = CancellationToken()

        def _work() -> dict[str, object]:
            result = validate_workbooks(
                template_id=ID_COURSE_SETUP,
                workbook_paths=staged_uploaded_paths,
                workbook_kind="course_details",
                cancel_token=token,
            )
            total = len(staged_uploaded_paths)
            result["total"] = total
            return result

        def _on_success(result: object) -> None:
            data = result if isinstance(result, dict) else {}
            valid_paths = [p for p in data.get("valid_paths", []) if isinstance(p, str) and p]
            invalid_paths = [p for p in data.get("invalid_paths", []) if isinstance(p, str) and p]
            mismatched_paths = [p for p in data.get("mismatched_paths", []) if isinstance(p, str) and p]
            duplicate_paths = [p for p in data.get("duplicate_paths", []) if isinstance(p, str) and p]
            duplicate_sections = [p for p in data.get("duplicate_sections", []) if isinstance(p, str) and p]
            rejection_items = [
                item for item in data.get("rejections", []) if isinstance(item, dict)
            ]
            total = int(data.get("total", 0))
            replacing_existing = had_previous_validated
            self._validated_course_details_paths = list(valid_paths)
            self._uploaded_course_details_paths = list(valid_paths)
            self.course_details_paths = list(valid_paths)
            self.marks_template_paths = []
            self.marks_template_path = None
            self._set_course_details_widget_files(self.course_details_paths)
            self._ready_for_marks_generation = bool(self.course_details_paths)

            if replacing_existing:
                self._publish_status_key("instructor.status.course_details_replaced")
            elif self.course_details_paths:
                self._publish_status_key("instructor.status.course_details_validated")
            self._publish_status_key(
                "instructor.status.course_details_validation_progress",
                valid=len(self.course_details_paths),
                total=total,
            )
            log_process_message(
                process_name,
                logger=self._logger,
                success_message=f"{process_name} completed successfully.",
                user_success_message=user_success_message,
            )

            duplicate_count = len(duplicate_paths) + len(duplicate_sections)
            self._publish_course_details_rejection_details(rejection_items)
            if invalid_paths or mismatched_paths or duplicate_count:
                self._runtime.notify_message_key(
                    "instructor.toast.course_details_validation_summary",
                    channels=("toast",),
                    kwargs={
                        "valid": len(self.course_details_paths),
                        "invalid": len(invalid_paths),
                        "mismatched": len(mismatched_paths),
                        "duplicates": duplicate_count,
                    },
                    toast_title_key="instructor.msg.validation_title",
                    toast_level="warning",
                )

        def _on_failure(exc: Exception) -> None:
            self._handle_async_failure(exc, process_name=process_name, user_error_message=user_error_message)

        self._start_async_operation(token=token, job_id=None, work=_work, on_success=_on_success, on_failure=_on_failure)

    def _prepare_marks_template_async(self) -> None:
        if self.state.busy:
            return
        source_paths = [path for path in self.course_details_paths if path]
        if not source_paths:
            self._runtime.notify_message_key(
                "instructor.validation.course_details_missing",
                channels=("toast",),
                toast_title_key="instructor.msg.validation_title",
                toast_level="warning",
            )
            return

        output_dir = QFileDialog.getExistingDirectory(
            self,
            t("instructor.dialog.marks_template.save_title"),
            resolve_dialog_start_path(APP_NAME),
        )
        if not output_dir:
            return
        self._remember_dialog_dir_safe(output_dir)

        process_key = "instructor.log.process.generate_marks_template"
        process_name = t(process_key)
        user_success_message, user_error_message = _localized_log_messages(process_key)
        template_id = ID_COURSE_SETUP

        all_generated_paths: list[str] = []
        per_file_native_limit = 5

        def _start_generation(
            run_sources: list[str],
            *,
            overwrite_existing: bool,
            output_path_overrides: dict[str, str] | None = None,
        ) -> None:
            token = CancellationToken()

            def _work() -> dict[str, object]:
                token.raise_if_cancelled()
                generation_context: dict[str, object] = {"overwrite_existing": overwrite_existing}
                if output_path_overrides:
                    generation_context["output_path_overrides"] = output_path_overrides
                return generate_workbooks(
                    template_id=template_id,
                    workbook_paths=run_sources,
                    output_dir=output_dir,
                    workbook_kind="marks_template",
                    cancel_token=token,
                    context=generation_context,
                )

            def _on_success(result: object) -> None:
                data = result if isinstance(result, dict) else {}
                generated = [
                    str(path)
                    for path in data.get("generated_workbook_paths", [])
                    if isinstance(path, str) and path.strip()
                ]
                for path in generated:
                    if path not in all_generated_paths:
                        all_generated_paths.append(path)

                overwrite_conflicts = extract_overwrite_conflicts_from_generation_result(data)
                if overwrite_conflicts and not overwrite_existing:
                    resolved = resolve_overwrite_conflicts(
                        overwrite_conflicts,
                        per_file_native_limit=per_file_native_limit,
                        ask_overwrite_all=self._prompt_overwrite_all_conflicts,
                        ask_output_path=self._prompt_output_path_for_collision,
                    )
                    retry_sources = resolved.retry_sources
                    retry_overrides = resolved.output_path_overrides
                    if retry_sources:
                        _start_generation(
                            retry_sources,
                            overwrite_existing=True,
                            output_path_overrides=retry_overrides,
                        )
                        return

                results = data.get("results", {})
                failed: list[dict[str, str]] = []
                if isinstance(results, dict):
                    for item in results.values():
                        if not isinstance(item, dict):
                            continue
                        if str(item.get("status") or "").strip() != "failed":
                            continue
                        source = str(item.get("source_path") or "").strip()
                        reason = str(item.get("reason") or "").strip()
                        failed.append({"source": source, "reason": reason or "failed"})
                total = int(data.get("total", 0))
                skipped = int(data.get("skipped", 0))

                self._publish_status_key(
                    "instructor.status.marks_template_generation_progress",
                    processed=total,
                    total=total,
                )

                self.marks_template_paths = list(all_generated_paths)
                self.marks_template_path = self.marks_template_paths[-1] if self.marks_template_paths else None
                if self.marks_template_paths:
                    self._publish_status_key("instructor.status.marks_template_generated")

                if failed:
                    details = "; ".join(
                        f"{item.get('source', '')} -> {item.get('reason', '')}" for item in failed[:5]
                    )
                    self._publish_status_key("instructor.status.marks_template_per_file_failures", details=details)

                log_process_message(
                    process_name,
                    logger=self._logger,
                    success_message=f"{process_name} completed successfully.",
                    user_success_message=user_success_message,
                )
                self._runtime.notify_message_key(
                    "instructor.toast.marks_template_generation_summary",
                    channels=("toast",),
                    kwargs={
                        "generated": len(self.marks_template_paths),
                        "processed": total,
                        "total": total,
                        "failed": len(failed),
                        "skipped": skipped,
                    },
                    toast_title_key="instructor.msg.success_title",
                    toast_level="success" if self.marks_template_paths else "warning",
                )

            def _on_failure(exc: Exception) -> None:
                self._handle_async_failure(exc, process_name=process_name, user_error_message=user_error_message)

            self._start_async_operation(
                token=token,
                job_id=None,
                work=_work,
                on_success=_on_success,
                on_failure=_on_failure,
            )

        _start_generation(source_paths, overwrite_existing=False)

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        self._runtime.remember_dialog_dir_safe(selected_path)

    def _merge_uploaded_paths(self, selected_paths: list[str]) -> list[str]:
        merged: list[str] = []
        seen: set[str] = set()
        for path in [*self._uploaded_course_details_paths, *selected_paths]:
            key = canonical_path_key(path)
            if key in seen:
                continue
            seen.add(key)
            merged.append(path)
        return merged

    def _set_course_details_widget_files(self, paths: list[str]) -> None:
        self._syncing_drop_widget_files = True
        try:
            self.course_details_drop_widget.set_files(paths)
        finally:
            self._syncing_drop_widget_files = False

    def _publish_course_details_rejection_details(self, rejection_items: list[dict[str, object]]) -> None:
        for item in rejection_items:
            issue_payload = item.get("issue")
            if not isinstance(issue_payload, dict):
                continue
            workbook = item.get("path")
            file_path = str(workbook).strip() if isinstance(workbook, str) else None
            self._runtime.notify_validation_issue(
                cast(dict[str, object], issue_payload),
                file_path=file_path,
                channels=("status", "activity_log"),
            )

    def _setup_ui_logging(self) -> None:
        self._runtime.setup_ui_logging()

    def _append_user_log(self, message: str) -> None:
        self._runtime.append_user_log(message)

    def _publish_status(self, message: str) -> None:
        self._runtime.notify_message(message, channels=("status", "activity_log"))

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        self._runtime.notify_message_key(text_key, channels=("status", "activity_log"), kwargs=kwargs)

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        host_window = self.window()
        set_switch = getattr(host_window, "set_language_switch_enabled", None)
        if callable(set_switch):
            set_switch(not busy)
        self._refresh_ui()

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work,
        on_success,
        on_failure,
        on_finally=None,
    ) -> None:
        self._runtime.set_async_runner(self._async_runner)
        self._runtime.start_async_operation(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
            on_finally=on_finally,
        )

    def _prompt_overwrite_all_conflicts(self, output_paths: list[str]) -> bool:
        preview = "\n".join(output_paths[:5])
        extra_count = len(output_paths) - 5
        extra_suffix = f"\n... (+{extra_count} more)" if extra_count > 0 else ""
        choice = QMessageBox.question(
            self,
            t("instructor.msg.validation_title"),
            (
                "Some marks-template outputs already exist.\n\n"
                f"{preview}{extra_suffix}\n\n"
                "Overwrite all collided files?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return choice == QMessageBox.StandardButton.Yes

    def _prompt_output_path_for_collision(self, suggested_output: str) -> str | None:
        selected_output, _ = QFileDialog.getSaveFileName(
            self,
            t("instructor.dialog.marks_template.save_title"),
            suggested_output,
            t("instructor.dialog.filter.excel"),
        )
        chosen_output = str(selected_output or "").strip()
        if not chosen_output:
            return None
        self._remember_dialog_dir_safe(chosen_output)
        return chosen_output

    def _handle_async_failure(self, exc: Exception, *, process_name: str, user_error_message: str) -> None:
        if isinstance(exc, JobCancelledError):
            return
        if isinstance(exc, ValidationError):
            self._runtime.notify_message(
                str(exc),
                channels=("toast",),
                toast_title=t("instructor.msg.validation_title"),
                toast_level="error",
            )
            log_process_message(
                process_name,
                logger=self._logger,
                error=exc,
                user_error_message=user_error_message,
            )
            return
        if isinstance(exc, AppSystemError):
            self._runtime.notify_message(
                str(exc),
                channels=("toast",),
                toast_title=t("instructor.msg.error_title"),
                toast_level="error",
            )
            log_process_message(
                process_name,
                logger=self._logger,
                error=exc,
                user_error_message=user_error_message,
            )
            return
        self._runtime.notify_message_key(
            "instructor.msg.failed_to_do",
            channels=("toast",),
            kwargs={"action": process_name},
            toast_title_key="instructor.msg.error_title",
            toast_level="error",
        )
        log_process_message(
            process_name,
            logger=self._logger,
            error=exc,
            user_error_message=user_error_message,
        )

    def closeEvent(self, event) -> None:
        self._is_closing = True
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)


def _messages_namespace() -> dict[str, object]:
    return dict(_default_messages_namespace(translate=t))


def _localized_log_messages(process_key: str) -> tuple[str, str]:
    return (
        _build_status_message(
            "instructor.log.completed_process",
            translate=t,
            kwargs={"process": {"__t_key__": process_key}},
            fallback=t("instructor.log.completed_process", process=t(process_key)),
        ),
        _build_status_message(
            "instructor.log.error_while_process",
            translate=t,
            kwargs={"process": {"__t_key__": process_key}},
            fallback=t("instructor.log.error_while_process", process=t(process_key)),
        ),
    )


__all__ = ["InstructorModule"]

