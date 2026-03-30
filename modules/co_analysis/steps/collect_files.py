"""CO Analysis step: validate and collect uploaded workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES,
    CO_ANALYSIS_WORKFLOW_STEP_ID_COLLECT_FILES,
    WORKFLOW_PAYLOAD_KEY_PATH,
)
from common.exceptions import AppSystemError
from common.jobs import CancellationToken, JobContext
from modules.co_analysis.steps.shared_execution import handle_step_failure

_MAX_ANOMALY_WARNING_LINES = 12


def _compact_context_text(context: object) -> str:
    if not isinstance(context, dict):
        return ""
    fields = ("sheet_name", "cell", "row", "range", "column")
    parts: list[str] = []
    for key in fields:
        value = context.get(key)
        token = str(value).strip() if value is not None else ""
        if token:
            parts.append(f"{key}={token}")
    return ", ".join(parts)


def _reason_key_counts(
    *,
    unsupported_or_missing_files: int,
    invalid_source_workbook_files: int,
    invalid_system_hash_files: int,
    invalid_marks_unfilled_files: int,
    invalid_layout_manifest_files: int,
    invalid_template_mismatch_files: int,
    invalid_mark_value_files: int,
    invalid_other_validation_files: int,
    duplicates: int,
    duplicate_reg_number_files: int,
    cohort_mismatch_files: int,
) -> list[tuple[str, int]]:
    return [
        ("co_analysis.status.ignored_reason.unsupported_or_missing", unsupported_or_missing_files),
        ("co_analysis.status.ignored_reason.invalid_workbook", invalid_source_workbook_files),
        ("co_analysis.status.ignored_reason.invalid_hash", invalid_system_hash_files),
        ("co_analysis.status.ignored_reason.marks_unfilled", invalid_marks_unfilled_files),
        ("co_analysis.status.ignored_reason.layout_manifest", invalid_layout_manifest_files),
        ("co_analysis.status.ignored_reason.template_mismatch", invalid_template_mismatch_files),
        ("co_analysis.status.ignored_reason.mark_value", invalid_mark_value_files),
        ("co_analysis.status.ignored_reason.invalid_other", invalid_other_validation_files),
        ("co_analysis.status.ignored_reason.duplicates", duplicates),
        ("co_analysis.status.ignored_reason.duplicate_reg", duplicate_reg_number_files),
        ("co_analysis.status.ignored_reason.cohort_mismatch", cohort_mismatch_files),
    ]


def _ignored_summary_kwargs(reason_key_by_count: list[tuple[str, int]]) -> dict[str, object]:
    parts = [
        {"__t_key__": key, "kwargs": {"count": value}}
        for key, value in reason_key_by_count
        if value > 0
    ]
    max_parts = 12
    payload: dict[str, object] = {}
    for idx in range(max_parts):
        key = f"part{idx + 1}"
        payload[key] = parts[idx] if idx < len(parts) else ""
    for idx in range(1, max_parts):
        left = payload.get(f"part{idx}")
        right = payload.get(f"part{idx + 1}")
        payload[f"sep{idx}_{idx + 1}"] = "; " if left and right else ""
    return payload


class _State(Protocol):
    busy: bool


class _WorkflowService(Protocol):
    def create_job_context(self, *, step_id: str, payload: Mapping[str, object] | None = None) -> JobContext:
        ...

    def collect_files(
        self,
        candidate_paths: list[str],
        *,
        existing_paths: list[str],
        validate_uploaded_source_workbook: Callable[[str | Path], None],
        consume_last_source_anomaly_warnings: Callable[[], list[str]],
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, object]:
        ...


class _Module(Protocol):
    state: _State
    _files: list[Path]
    _pending_drop_batches: list[list[str]]
    _logger: object
    _workflow_service: _WorkflowService

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        ...

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
        on_finally: Callable[[], None] | None = None,
    ) -> None:
        ...

    def _drain_next_batch(self) -> None:
        ...

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        ...

    def _refresh_ui(self) -> None:
        ...


class _Ns(TypedDict):
    _validate_uploaded_source_workbook: Callable[[str | Path], None]
    _consume_last_source_anomaly_warnings: Callable[[], list[str]]
    t: Callable[..., str]
    show_toast: Callable[..., None]
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    JobCancelledError: type[Exception]


def collect_files_async(module: object, candidate_paths: list[str], *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_Module, module)
    typed_ns = cast(_Ns, ns)
    if not candidate_paths:
        return
    if typed_module.state.busy:
        typed_module._pending_drop_batches.append(candidate_paths)
        typed_module._publish_status_key("coordinator.status.queued", count=len(candidate_paths))
        return

    process_name = CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES
    token = CancellationToken()
    job_context = typed_module._workflow_service.create_job_context(
        step_id=CO_ANALYSIS_WORKFLOW_STEP_ID_COLLECT_FILES,
        payload={WORKFLOW_PAYLOAD_KEY_PATH: list(candidate_paths)},
    )
    typed_module._publish_status_key("coordinator.status.processing_started")

    def _on_success(result: object) -> None:
        if not isinstance(result, dict):
            raise AppSystemError("CO Analysis collect_files returned unexpected result.")
        added_paths = [Path(str(value)) for value in cast(list[object], result.get("added", []))]
        duplicates = int(result.get("duplicates", 0))
        invalid = int(result.get("invalid", 0))
        ignored = int(result.get("ignored", duplicates + invalid))
        unsupported_or_missing_files = int(result.get("unsupported_or_missing_files", 0))
        invalid_source_workbook_files = int(result.get("invalid_source_workbook_files", 0))
        duplicate_reg_number_files = int(result.get("duplicate_reg_number_files", 0))
        cohort_mismatch_files = int(result.get("cohort_mismatch_files", result.get("co_count_mismatch_files", 0)))
        invalid_system_hash_files = int(result.get("invalid_system_hash_files", 0))
        invalid_marks_unfilled_files = int(result.get("invalid_marks_unfilled_files", 0))
        invalid_layout_manifest_files = int(result.get("invalid_layout_manifest_files", 0))
        invalid_template_mismatch_files = int(result.get("invalid_template_mismatch_files", 0))
        invalid_mark_value_files = int(result.get("invalid_mark_value_files", 0))
        invalid_other_validation_files = int(result.get("invalid_other_validation_files", 0))
        validation_failures = cast(list[object], result.get("validation_failures", []))
        anomaly_warnings = [str(item) for item in cast(list[object], result.get("anomaly_warnings", [])) if str(item)]
        reason_key_by_count = _reason_key_counts(
            unsupported_or_missing_files=unsupported_or_missing_files,
            invalid_source_workbook_files=invalid_source_workbook_files,
            invalid_system_hash_files=invalid_system_hash_files,
            invalid_marks_unfilled_files=invalid_marks_unfilled_files,
            invalid_layout_manifest_files=invalid_layout_manifest_files,
            invalid_template_mismatch_files=invalid_template_mismatch_files,
            invalid_mark_value_files=invalid_mark_value_files,
            invalid_other_validation_files=invalid_other_validation_files,
            duplicates=duplicates,
            duplicate_reg_number_files=duplicate_reg_number_files,
            cohort_mismatch_files=cohort_mismatch_files,
        )

        typed_module._add_uploaded_paths(added_paths)
        if added_paths:
            typed_module._publish_status_key(
                "coordinator.status.added",
                added=len(added_paths),
                total=len(typed_module._files),
            )
        if duplicates or invalid:
            toast_lines = [typed_ns["t"](key, count=value) for key, value in reason_key_by_count if value > 0]
            typed_ns["show_toast"](
                typed_module,
                "\n".join(toast_lines)
                if toast_lines
                else typed_ns["t"](
                    "co_analysis.validation.rejection_breakdown_body",
                    unsupported_or_missing=unsupported_or_missing_files,
                    invalid_workbook=invalid_source_workbook_files,
                    duplicates=duplicates,
                    duplicate_reg=duplicate_reg_number_files,
                    cohort_mismatch=cohort_mismatch_files,
                    invalid_hash=invalid_system_hash_files,
                    marks_unfilled=invalid_marks_unfilled_files,
                    layout_manifest=invalid_layout_manifest_files,
                    template_mismatch=invalid_template_mismatch_files,
                    mark_value=invalid_mark_value_files,
                    invalid_other=invalid_other_validation_files,
                ),
                title=typed_ns["t"]("co_analysis.validation.rejection_breakdown_title"),
                level="warning",
            )
        if duplicate_reg_number_files:
            typed_module._publish_status_key(
                "co_analysis.status.duplicate_reg_numbers",
                count=duplicate_reg_number_files,
            )
            typed_ns["show_toast"](
                typed_module,
                typed_ns["t"]("coordinator.regno_dedup.body", count=duplicate_reg_number_files),
                title=typed_ns["t"]("coordinator.regno_dedup.title"),
                level="warning",
            )
        if cohort_mismatch_files:
            typed_module._publish_status_key(
                "co_analysis.status.cohort_mismatch",
                count=cohort_mismatch_files,
            )
            typed_ns["show_toast"](
                typed_module,
                typed_ns["t"]("co_analysis.validation.cohort_mismatch_body", count=cohort_mismatch_files),
                title=typed_ns["t"]("coordinator.title"),
                level="warning",
            )
        if ignored:
            summary_kwargs = _ignored_summary_kwargs(reason_key_by_count)
            typed_module._publish_status_key(
                "co_analysis.status.ignored_summary",
                count=ignored,
                **summary_kwargs,
            )
        if validation_failures:
            for item in validation_failures[:12]:
                if not isinstance(item, dict):
                    continue
                file_name = str(item.get("file", "")).strip()
                code = str(item.get("code", "")).strip()
                context = item.get("context", {})
                typed_module._publish_status_key(
                    "co_analysis.status.rejected_code_with_context",
                    file=file_name or "-",
                    code=code or "UNKNOWN",
                    context=_compact_context_text(context),
                )
        if anomaly_warnings:
            typed_module._publish_status_key(
                "co_analysis.status.validation_warnings",
                count=len(anomaly_warnings),
            )
            displayed_warnings = anomaly_warnings[:_MAX_ANOMALY_WARNING_LINES]
            for warning in displayed_warnings:
                typed_module._publish_status_key(
                    "co_analysis.status.validation_warning_line",
                    warning=warning,
                )
            hidden_warning_count = len(anomaly_warnings) - len(displayed_warnings)
            if hidden_warning_count > 0:
                typed_module._publish_status_key(
                    "co_analysis.status.validation_warning_more",
                    count=hidden_warning_count,
                )
            typed_ns["show_toast"](
                typed_module,
                typed_ns["t"]("co_analysis.validation.anomaly_warnings_body", count=len(anomaly_warnings)),
                title=typed_ns["t"]("coordinator.title"),
                level="warning",
            )
        typed_ns["log_process_message"](
            process_name,
            logger=typed_module._logger,
            success_message=(
                "collecting co analysis files completed successfully. "
                f"added={len(added_paths)}, duplicates={duplicates}, invalid={invalid}, "
                f"cohort_mismatch_files={cohort_mismatch_files}, "
                f"ignored={ignored}, duplicate_reg_number_files={duplicate_reg_number_files}, "
                f"unsupported_or_missing_files={unsupported_or_missing_files}, "
                f"invalid_source_workbook_files={invalid_source_workbook_files}, "
                f"invalid_system_hash_files={invalid_system_hash_files}, "
                f"invalid_marks_unfilled_files={invalid_marks_unfilled_files}, "
                f"invalid_layout_manifest_files={invalid_layout_manifest_files}, "
                f"invalid_template_mismatch_files={invalid_template_mismatch_files}, "
                f"invalid_mark_value_files={invalid_mark_value_files}, "
                f"invalid_other_validation_files={invalid_other_validation_files}, "
                f"anomaly_warnings={len(anomaly_warnings)}"
            ),
            user_success_message=typed_ns["build_i18n_log_message"](
                "coordinator.status.processing_completed",
                fallback=typed_ns["t"]("coordinator.status.processing_completed"),
            ),
            job_id=job_context.job_id,
            step_id=job_context.step_id,
        )
        typed_module._refresh_ui()

    def _on_failure(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            job_id=job_context.job_id,
            step_id=job_context.step_id,
        )

    typed_module._start_async_operation(
        token=token,
        job_id=job_context.job_id,
        work=lambda: typed_module._workflow_service.collect_files(
            list(candidate_paths),
            existing_paths=[str(path) for path in typed_module._files],
            validate_uploaded_source_workbook=typed_ns["_validate_uploaded_source_workbook"],
            consume_last_source_anomaly_warnings=typed_ns["_consume_last_source_anomaly_warnings"],
            context=job_context,
            cancel_token=token,
        ),
        on_success=_on_success,
        on_failure=_on_failure,
        on_finally=typed_module._drain_next_batch,
    )



