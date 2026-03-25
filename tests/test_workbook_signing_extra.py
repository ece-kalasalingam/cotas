from common.workbook_integrity import workbook_signing as ws


def test_verify_payload_signature_empty_and_version_mismatch(monkeypatch) -> None:
    monkeypatch.setattr(ws, "ensure_workbook_secret_policy", lambda: None)
    monkeypatch.setattr(ws, "get_workbook_password", lambda: "secret")
    monkeypatch.setattr(ws, "_accepted_secrets", lambda: ("secret",))
    monkeypatch.setattr(ws, "WORKBOOK_SIGNATURE_VERSION", "v1")

    assert ws.verify_payload_signature("x", "") is False
    assert ws.verify_payload_signature("x", "v2:abcd") is False


def test_verify_payload_signature_hmac_branch_no_secret_match(monkeypatch) -> None:
    monkeypatch.setattr(ws, "ensure_workbook_secret_policy", lambda: None)
    monkeypatch.setattr(ws, "_accepted_secrets", lambda: ("a", "b"))
    monkeypatch.setattr(ws, "WORKBOOK_SIGNATURE_VERSION", "v1")

    assert ws.verify_payload_signature("payload", "v1:not-a-valid-digest") is False

