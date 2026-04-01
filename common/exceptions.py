"""Application exception hierarchy.

This module defines typed exceptions used across the project so callers can
handle expected failures (validation/configuration) differently from unexpected
system failures.
"""

from __future__ import annotations

from typing import Any


class AppError(Exception):
    """Base class for all application-specific exceptions."""


class ValidationError(AppError):
    """Raised when user or business-rule validation fails.

    `code` and `context` allow UI layers to map to localized messages while
    keeping processing logic language-neutral.
    """

    def __init__(
        self,
        message: str = "",
        *,
        code: str = "VALIDATION_ERROR",
        context: dict[str, Any] | None = None,
    ) -> None:
        """Init.
        
        Args:
            message: Parameter value (str).
            code: Parameter value (str).
            context: Parameter value (dict[str, Any] | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        super().__init__(message or code)
        self.code = code
        self.context = dict(context or {})


class ConfigurationError(AppError):
    """Raised when static configuration or policy constants are invalid."""


class AppSystemError(AppError):
    """Raised for unexpected internal system-level failures."""


class JobCancelledError(AppError):
    """Raised when a cancellable job is interrupted by user/system request."""
