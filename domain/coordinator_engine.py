"""Template-routing facade for coordinator processing workflows."""

from __future__ import annotations

from domain.template_versions import course_setup_v1_coordinator_engine as _impl

# COURSE_SETUP_V1 implementation is centralized under template_versions.
# Re-export symbols here so existing module/service callsites remain stable.
__all__ = [name for name in dir(_impl) if not name.startswith("__")]
globals().update({name: getattr(_impl, name) for name in __all__})
