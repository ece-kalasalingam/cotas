from __future__ import annotations

import os


# Keep tests deterministic and aligned with production config contract.
os.environ.setdefault("FOCUS_WORKBOOK_PASSWORD", "test-workbook-password-2026")
