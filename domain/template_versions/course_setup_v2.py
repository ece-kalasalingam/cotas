"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.utils import normalize
from domain.template_versions.course_setup_v2_impl import strategy_bindings as _bindings

_SUPPORTED_OPERATIONS = frozenset(
    {
        "generate_workbook",
        "generate_workbooks",
        "validate_workbooks",
        "consume_last_marks_anomaly_warnings",
        "extract_course_metadata_and_students",
    }
)
_SUPPORTED_OPERATION_TOKENS = frozenset(normalize(value) for value in _SUPPORTED_OPERATIONS)


@dataclass(slots=True, frozen=True)
class CourseSetupV2Strategy:
    template_id: str = "COURSE_SETUP_V2"

    def supports_operation(self, operation: str) -> bool:
        """Supports operation.
        
        Args:
            operation: Parameter value (str).
        
        Returns:
            bool: Return value.
        
        Raises:
            None.
        """
        return normalize(operation) in _SUPPORTED_OPERATION_TOKENS

    def generate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        output_path: str | Path,
        workbook_name: str | None,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> object:
        """Generate workbook.
        
        Args:
            template_id: Parameter value (str).
            workbook_kind: Parameter value (str).
            output_path: Parameter value (str | Path).
            workbook_name: Parameter value (str | None).
            cancel_token: Parameter value (CancellationToken | None).
            context: Parameter value (Mapping[str, Any] | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(actual_template_id=template_id, expected_template_id=self.template_id)
        resolved_workbook_name = (workbook_name or Path(output_path).name).strip()
        if not resolved_workbook_name:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_NAME_REQUIRED",
            )
        kind = normalize(workbook_kind)
        if kind == "course_details_template":
            return _bindings.course_template_generator()(
                output_path=output_path,
                cancel_token=cancel_token,
            )
        if kind == "co_description_template":
            return _bindings.co_description_template_generator()(
                output_path=output_path,
                cancel_token=cancel_token,
            )
        if kind == "co_attainment":
            inputs = _bindings.co_attainment_generation_inputs(
                context=context,
                output_path=output_path,
                default_template_id=self.template_id,
            )
            return _bindings.co_attainment_generator()(
                source_paths=inputs["source_paths"],
                output_path=inputs["output_path"],
                token=cancel_token or CancellationToken(),
                total_outcomes=inputs["total_outcomes"],
                template_id=inputs["template_id"],
                thresholds=inputs["thresholds"],
                co_attainment_percent=inputs["co_attainment_percent"],
                co_attainment_level=inputs["co_attainment_level"],
                generate_word_report=bool(inputs.get("generate_word_report", False)),
                word_output_path=inputs.get("word_output_path"),
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def validate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> dict[str, object]:
        """Validate workbooks.
        
        Args:
            template_id: Parameter value (str).
            workbook_kind: Parameter value (str).
            workbook_paths: Parameter value (Sequence[str | Path]).
            cancel_token: Parameter value (CancellationToken | None).
            context: Parameter value (Mapping[str, Any] | None).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        from domain.template_strategy_router import assert_template_id_matches

        del context
        assert_template_id_matches(actual_template_id=template_id, expected_template_id=self.template_id)
        kind = normalize(workbook_kind)
        if kind == "course_details":
            return _bindings.course_template_batch_validator()(
                workbook_paths=workbook_paths,
                cancel_token=cancel_token,
            )
        if kind == "marks_template":
            return _bindings.marks_template_batch_validator()(
                workbook_paths=workbook_paths,
                template_id=self.template_id,
                cancel_token=cancel_token,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def generate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        output_dir: str | Path,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> dict[str, object]:
        """Generate workbooks.
        
        Args:
            template_id: Parameter value (str).
            workbook_kind: Parameter value (str).
            workbook_paths: Parameter value (Sequence[str | Path]).
            output_dir: Parameter value (str | Path).
            cancel_token: Parameter value (CancellationToken | None).
            context: Parameter value (Mapping[str, Any] | None).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(actual_template_id=template_id, expected_template_id=self.template_id)
        kind = normalize(workbook_kind)
        if kind == "marks_template":
            return _bindings.marks_template_batch_generator()(
                workbook_paths=workbook_paths,
                output_dir=Path(output_dir),
                allow_overwrite=_bindings.overwrite_existing_enabled(context),
                output_path_overrides=_bindings.output_path_overrides_from_context(context),
                cancel_token=cancel_token,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def consume_last_marks_anomaly_warnings(self) -> list[str]:
        """Consume last marks anomaly warnings.
        
        Args:
            None.
        
        Returns:
            list[str]: Return value.
        
        Raises:
            None.
        """
        return _bindings.consume_last_marks_anomaly_warnings()

    def extract_course_metadata_and_students(
        self,
        workbook_path: str | Path,
        *,
        template_id: str,
    ) -> tuple[set[str], dict[str, str]]:
        """Extract course metadata and students.
        
        Args:
            workbook_path: Parameter value (str | Path).
            template_id: Parameter value (str).
        
        Returns:
            tuple[set[str], dict[str, str]]: Return value.
        
        Raises:
            None.
        """
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(actual_template_id=template_id, expected_template_id=self.template_id)
        return _bindings.course_metadata_students_extractor()(Path(workbook_path))


__all__ = ["CourseSetupV2Strategy"]
