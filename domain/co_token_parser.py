"""Compatibility shim forwarding CO token parsing to COURSE_SETUP_V2 implementation."""

from __future__ import annotations

from domain.template_versions.course_setup_v2_impl.co_token_parser import parse_co_tokens

__all__ = ["parse_co_tokens"]

