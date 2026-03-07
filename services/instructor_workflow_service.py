"""Instructor workflow service orchestration."""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
import time
from inspect import Signature, signature
from pathlib import Path
from typing import Any, Callable, Mapping, TypeVar

from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, JobContext
from modules.instructor import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)

_logger = logging.getLogger(__name__)
_T = TypeVar("_T")


class InstructorWorkflowService:
    def create_job_context(self, *, step_id: str, payload: Mapping[str, Any] | None = None) -> JobContext:
        return JobContext.create(step_id=step_id, payload=payload)

    def generate_course_details_template(
        self,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation="generate_course_details_template",
            cancel_token=cancel_token,
            work=lambda: self._call_with_optional_cancel_token(
                generate_course_details_template,
                output_path,
                cancel_token=cancel_token,
            ),
        )

    def validate_course_details_workbook(
        self,
        workbook_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> str:
        return self._execute_with_telemetry(
            context=context,
            operation="validate_course_details_workbook",
            cancel_token=cancel_token,
            work=lambda: validate_course_details_workbook(workbook_path),
        )

    def generate_marks_template(
        self,
        course_details_path: str | Path,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation="generate_marks_template",
            cancel_token=cancel_token,
            work=lambda: self._call_with_optional_cancel_token(
                generate_marks_template_from_course_details,
                course_details_path,
                output_path,
                cancel_token=cancel_token,
            ),
        )

    def generate_final_report(
        self,
        filled_marks_path: str | Path,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation="generate_final_report",
            cancel_token=cancel_token,
            work=lambda: self._atomic_copy_file(filled_marks_path, output_path),
        )

    @staticmethod
    def _raise_if_cancelled(cancel_token: CancellationToken | None) -> None:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()

    def _execute_with_telemetry(
        self,
        *,
        context: JobContext,
        operation: str,
        cancel_token: CancellationToken | None,
        work: Callable[[], _T],
    ) -> _T:
        started_at = time.perf_counter()
        _logger.info(
            "Instructor workflow step started.",
            extra={
                "user_message": f"{operation} started.",
                "job_id": context.job_id,
                "step_id": context.step_id,
            },
        )
        try:
            self._raise_if_cancelled(cancel_token)
            result = work()
            self._raise_if_cancelled(cancel_token)
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _logger.info(
                "Instructor workflow step completed.",
                extra={
                    "user_message": f"{operation} completed in {duration_ms} ms.",
                    "job_id": context.job_id,
                    "step_id": context.step_id,
                },
            )
            return result
        except JobCancelledError:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _logger.info(
                "Instructor workflow step cancelled.",
                extra={
                    "user_message": f"{operation} cancelled after {duration_ms} ms.",
                    "job_id": context.job_id,
                    "step_id": context.step_id,
                },
            )
            raise
        except Exception as exc:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _logger.exception(
                "Instructor workflow step failed.",
                exc_info=exc,
                extra={
                    "user_message": f"{operation} failed after {duration_ms} ms.",
                    "job_id": context.job_id,
                    "step_id": context.step_id,
                },
            )
            raise

    @staticmethod
    def _atomic_copy_file(source_path: str | Path, output_path: str | Path) -> Path:
        source = Path(source_path)
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        temp_name = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                delete=False,
                dir=str(output.parent),
                prefix=f"{output.name}.",
                suffix=".tmp",
            ) as temp_file:
                temp_name = temp_file.name
            shutil.copyfile(str(source), temp_name)
            os.replace(temp_name, output)
        except Exception:
            if temp_name:
                try:
                    Path(temp_name).unlink(missing_ok=True)
                except OSError:
                    pass
            raise
        return output

    @staticmethod
    def _call_with_optional_cancel_token(fn, *args: object, cancel_token: CancellationToken | None):
        try:
            fn_signature: Signature = signature(fn)
        except (TypeError, ValueError):
            fn_signature = Signature()
        if "cancel_token" in fn_signature.parameters:
            return fn(*args, cancel_token=cancel_token)
        return fn(*args)
