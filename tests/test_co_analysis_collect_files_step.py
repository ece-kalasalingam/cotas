from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from common.jobs import JobContext
from modules.co_analysis.steps.collect_files import collect_files_async


@dataclass
class _State:
    busy: bool = False


class _WorkflowService:
    def create_job_context(self, *, step_id: str, payload=None) -> JobContext:
        del payload
        return JobContext(job_id="job-1", step_id=step_id, language="en", created_at_utc="2026-01-01T00:00:00+00:00", payload={})

    def collect_files(
        self,
        candidate_paths,
        *,
        existing_paths,
        validate_uploaded_source_workbook,
        consume_last_source_anomaly_warnings,
        context,
        cancel_token=None,
    ):
        del candidate_paths
        del existing_paths
        del validate_uploaded_source_workbook
        del consume_last_source_anomaly_warnings
        del context
        del cancel_token
        return {
            "added": [],
            "duplicates": 0,
            "invalid": 0,
            "ignored": 0,
            "unsupported_or_missing_files": 0,
            "invalid_source_workbook_files": 0,
            "duplicate_reg_number_files": 0,
            "co_count_mismatch_files": 0,
            "invalid_system_hash_files": 0,
            "invalid_marks_unfilled_files": 0,
            "invalid_layout_manifest_files": 0,
            "invalid_template_mismatch_files": 0,
            "invalid_mark_value_files": 0,
            "invalid_other_validation_files": 0,
            "validation_failures": [],
            "anomaly_warnings": [f"file_{idx}.xlsx -> warning {idx}" for idx in range(15)],
        }


class _Module:
    def __init__(self) -> None:
        self.state = _State(False)
        self._files: list[Path] = []
        self._pending_drop_batches: list[list[str]] = []
        self._logger = object()
        self._workflow_service = _WorkflowService()
        self.published: list[tuple[str, dict[str, object]]] = []

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        self.published.append((text_key, dict(kwargs)))

    def _start_async_operation(
        self,
        *,
        token,
        job_id,
        work,
        on_success,
        on_failure,
        on_finally=None,
    ) -> None:
        del token
        del job_id
        try:
            on_success(work())
        except Exception as exc:  # pragma: no cover - defensive
            on_failure(exc)
        if on_finally is not None:
            on_finally()

    def _drain_next_batch(self) -> None:
        return

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        self._files.extend(added_paths)

    def _refresh_ui(self) -> None:
        return


def test_collect_files_async_emits_capped_anomaly_warning_lines() -> None:
    module = _Module()
    toasts: list[tuple[str, str, str]] = []

    def _show_toast(_module, body: str, *, title: str, level: str) -> None:
        toasts.append((body, title, level))

    ns = {
        "_validate_uploaded_source_workbook": lambda _path: None,
        "_consume_last_source_anomaly_warnings": lambda: [],
        "t": lambda key, **kwargs: key.format(**kwargs) if kwargs else key,
        "show_toast": _show_toast,
        "log_process_message": lambda *args, **kwargs: None,
        "build_i18n_log_message": lambda *args, **kwargs: "done",
        "JobCancelledError": Exception,
    }

    collect_files_async(module, ["a.xlsx"], ns=ns)

    warning_summary = [entry for entry in module.published if entry[0] == "co_analysis.status.validation_warnings"]
    warning_lines = [entry for entry in module.published if entry[0] == "co_analysis.status.validation_warning_line"]
    warning_more = [entry for entry in module.published if entry[0] == "co_analysis.status.validation_warning_more"]

    assert warning_summary and warning_summary[-1][1]["count"] == 15
    assert len(warning_lines) == 12
    assert warning_more and warning_more[-1][1]["count"] == 3
    assert any(level == "warning" for _body, _title, level in toasts)


def test_collect_files_async_uses_cohort_mismatch_reason_keys() -> None:
    class _CohortMismatchWorkflowService(_WorkflowService):
        def collect_files(
            self,
            candidate_paths,
            *,
            existing_paths,
            validate_uploaded_source_workbook,
            consume_last_source_anomaly_warnings,
            context,
            cancel_token=None,
        ):
            del candidate_paths
            del existing_paths
            del validate_uploaded_source_workbook
            del consume_last_source_anomaly_warnings
            del context
            del cancel_token
            return {
                "added": [],
                "duplicates": 0,
                "invalid": 1,
                "ignored": 1,
                "unsupported_or_missing_files": 0,
                "invalid_source_workbook_files": 0,
                "duplicate_reg_number_files": 0,
                "cohort_mismatch_files": 1,
                "invalid_system_hash_files": 0,
                "invalid_marks_unfilled_files": 0,
                "invalid_layout_manifest_files": 0,
                "invalid_template_mismatch_files": 0,
                "invalid_mark_value_files": 0,
                "invalid_other_validation_files": 0,
                "validation_failures": [],
                "anomaly_warnings": [],
            }

    module = _Module()
    module._workflow_service = _CohortMismatchWorkflowService()
    toasts: list[tuple[str, str, str]] = []

    def _show_toast(_module, body: str, *, title: str, level: str) -> None:
        toasts.append((body, title, level))

    ns = {
        "_validate_uploaded_source_workbook": lambda _path: None,
        "_consume_last_source_anomaly_warnings": lambda: [],
        "t": lambda key, **kwargs: key.format(**kwargs) if kwargs else key,
        "show_toast": _show_toast,
        "log_process_message": lambda *args, **kwargs: None,
        "build_i18n_log_message": lambda *args, **kwargs: "done",
        "JobCancelledError": Exception,
    }

    collect_files_async(module, ["a.xlsx"], ns=ns)

    assert any(key == "co_analysis.status.cohort_mismatch" for key, _ in module.published)
    assert any("co_analysis.status.ignored_reason.cohort_mismatch" in body for body, _title, _level in toasts)
