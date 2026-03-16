"""Workbook signature helpers."""

from __future__ import annotations

import hmac
from hashlib import sha256

from common.workbook_secret import (
    WORKBOOK_SIGNATURE_VERSION,
    ensure_workbook_secret_policy,
    get_workbook_password,
)

_SIGNATURE_DELIMITER = ":"


def sign_payload(payload: str) -> str:
    ensure_workbook_secret_policy()
    digest = _hmac_digest(payload, get_workbook_password())
    return f"{WORKBOOK_SIGNATURE_VERSION}{_SIGNATURE_DELIMITER}{digest}"


def verify_payload_signature(payload: str, signature: str) -> bool:
    ensure_workbook_secret_policy()
    token = str(signature or "").strip()
    if not token:
        return False

    if _SIGNATURE_DELIMITER in token:
        version, digest = token.split(_SIGNATURE_DELIMITER, 1)
        if version.strip().lower() != WORKBOOK_SIGNATURE_VERSION:
            return False
        for secret in _accepted_secrets():
            if hmac.compare_digest(digest, _hmac_digest(payload, secret)):
                return True
        return False

    # Backward compatibility for legacy signatures that were plain sha256 hex.
    for secret in _accepted_secrets():
        legacy = sha256(f"{payload}|{secret}".encode("utf-8")).hexdigest()
        if hmac.compare_digest(token, legacy):
            return True
    return False


def _accepted_secrets() -> tuple[str, ...]:
    return (get_workbook_password(),)


def _hmac_digest(payload: str, secret: str) -> str:
    return hmac.new(secret.encode("utf-8"), payload.encode("utf-8"), sha256).hexdigest()

