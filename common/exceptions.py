"""Application exception hierarchy.

This module defines typed exceptions used across the project so callers can
handle expected failures (validation/configuration) differently from unexpected
system failures.
"""

from __future__ import annotations


class AppError(Exception):
    """Base class for all application-specific exceptions."""


class ValidationError(AppError):
    """Raised when user or business-rule validation fails."""


class ConfigurationError(AppError):
    """Raised when static configuration or policy constants are invalid."""


class AppSystemError(AppError):
    """Raised for unexpected internal system-level failures."""
