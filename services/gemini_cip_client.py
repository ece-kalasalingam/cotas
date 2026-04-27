"""EXE-side HTTP client for the Cloudflare CIP Worker.

Reads cip_config.json from the app directory (alongside the EXE when frozen,
or the project root in dev). If the file is absent CIP generation is silently
disabled — the Word report is still produced without the CIP section.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from common.exceptions import AppError

CIP_CONFIG_FILE_NAME = "cip_config.json"
_REQUEST_TIMEOUT_SECONDS = 30


class CipWorkerError(AppError):
    """Raised when the CIP Worker call fails (network, HTTP, or bad response)."""

    def __init__(self, message: str, *, code: str = "CIP_WORKER_ERROR") -> None:
        super().__init__(message)
        self.code = code


def _resolve_config_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / CIP_CONFIG_FILE_NAME
    return Path(__file__).resolve().parent.parent / CIP_CONFIG_FILE_NAME


def load_cip_config(config_path: Path | str | None = None) -> dict[str, str] | None:
    """Load Worker URL and app token from cip_config.json.

    Returns None if the file is absent or incomplete — callers treat that as
    CIP disabled, not an error.
    """
    path = Path(config_path) if config_path else _resolve_config_path()
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise CipWorkerError(
            f"Invalid cip_config.json: {exc}",
            code="CIP_CONFIG_INVALID",
        ) from exc
    url = str(data.get("worker_url", "")).strip()
    token = str(data.get("app_token", "")).strip()
    if not url or not token:
        raise CipWorkerError(
            "cip_config.json is missing 'worker_url' or 'app_token'",
            code="CIP_CONFIG_INCOMPLETE",
        )
    return {"worker_url": url, "app_token": token}


def call_cip_worker(
    payload: dict[str, Any],
    *,
    config_path: Path | str | None = None,
    timeout: int = _REQUEST_TIMEOUT_SECONDS,
) -> str | None:
    """POST the compact CIP payload to the Cloudflare Worker.

    Returns the CIP report text on success, or None if cip_config.json is
    absent (CIP generation disabled). Raises CipWorkerError on failure.
    """
    config = load_cip_config(config_path)
    if config is None:
        return None

    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = Request(
        url=config["worker_url"],
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "X-App-Token": config["app_token"],
            "User-Agent": "COTAS-CIP-Client/1.0",
        }
    )

    try:
        with urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except HTTPError as exc:
        snippet = exc.read().decode("utf-8", errors="replace")[:200]
        raise CipWorkerError(
            f"Worker returned HTTP {exc.code}: {snippet}",
            code="CIP_WORKER_HTTP_ERROR",
        ) from exc
    except (TimeoutError, URLError, OSError) as exc:
        raise CipWorkerError(
            f"Network error calling CIP Worker: {exc}",
            code="CIP_WORKER_NETWORK_ERROR",
        ) from exc

    try:
        data = json.loads(raw)
    except Exception as exc:
        raise CipWorkerError(
            "CIP Worker returned non-JSON response",
            code="CIP_WORKER_INVALID_RESPONSE",
        ) from exc

    report_text = data.get("report_text")
    if not isinstance(report_text, str) or not report_text.strip():
        detail = str(data.get("error", "")).strip()
        raise CipWorkerError(
            f"CIP Worker returned no report_text{': ' + detail if detail else ''}",
            code="CIP_WORKER_EMPTY_RESPONSE",
        )

    return report_text.strip()


__all__ = ["CipWorkerError", "call_cip_worker", "load_cip_config"]
