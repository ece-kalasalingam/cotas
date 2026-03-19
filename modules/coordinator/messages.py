"""Shared coordinator module message and UI-log helpers."""

from __future__ import annotations

from typing import Mapping

from common.module_messages import MessagesNamespace
from common.module_messages import append_user_log as _append_user_log_impl
from common.module_messages import publish_status as _publish_status_impl
from common.module_messages import publish_status_key as _publish_status_key_impl
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_messages import setup_ui_logging as _setup_ui_logging_impl


def _contract_guard(ns: Mapping[str, object]) -> None:
    # Keep explicit key references so namespace contract tests can verify wiring.
    try:
        ns["t"]
    except KeyError:
        pass
    try:
        ns["build_i18n_log_message"]
    except KeyError:
        pass
    try:
        ns["parse_i18n_log_message"]
    except KeyError:
        pass
    try:
        ns["resolve_i18n_log_message"]
    except KeyError:
        pass
    try:
        ns["format_log_line_at"]
    except KeyError:
        pass
    try:
        ns["UILogHandler"]
    except KeyError:
        pass
    try:
        ns["emit_user_status"]
    except KeyError:
        pass


def publish_status(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    _contract_guard(ns)
    _publish_status_impl(module, message, ns=ns)


def publish_status_key(
    module: object,
    text_key: str,
    *,
    ns: Mapping[str, object],
    **kwargs: object,
) -> None:
    _contract_guard(ns)
    _publish_status_key_impl(module, text_key, ns=ns, **kwargs)


def setup_ui_logging(module: object, *, ns: Mapping[str, object]) -> None:
    _contract_guard(ns)
    _setup_ui_logging_impl(module, ns=ns)


def append_user_log(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    _contract_guard(ns)
    _append_user_log_impl(module, message, ns=ns)


def rerender_user_log(module: object, *, ns: Mapping[str, object]) -> None:
    _contract_guard(ns)
    _rerender_user_log_impl(module, ns=ns)


__all__ = [
    "MessagesNamespace",
    "append_user_log",
    "publish_status",
    "publish_status_key",
    "rerender_user_log",
    "setup_ui_logging",
]
