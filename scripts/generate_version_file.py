"""Generate PyInstaller version.txt from common/constants.py."""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from common.constants import (
    APP_EXECUTABLE_NAME,
    APP_INTERNAL_NAME,
    APP_PRODUCT_NAME,
    SYSTEM_VERSION,
)


def _parse_version(version: str) -> tuple[int, int, int, int]:
    parts = [p.strip() for p in version.split(".") if p.strip()]
    if not parts:
        raise ValueError("SYSTEM_VERSION is empty.")

    if len(parts) > 4:
        raise ValueError("SYSTEM_VERSION must have at most 4 numeric parts.")

    try:
        numeric = [int(p) for p in parts]
    except ValueError as exc:
        raise ValueError("SYSTEM_VERSION must contain only numeric parts.") from exc

    while len(numeric) < 4:
        numeric.append(0)
    return tuple(numeric)  # type: ignore[return-value]


def build_version_text() -> str:
    v1, v2, v3, v4 = _parse_version(SYSTEM_VERSION)
    version_str = f"{v1}.{v2}.{v3}.{v4}"

    return f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({v1}, {v2}, {v3}, {v4}),
    prodvers=({v1}, {v2}, {v3}, {v4}),
    mask=0x3F,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          '040904B0',
          [
            StringStruct('CompanyName', '{APP_PRODUCT_NAME}'),
            StringStruct('FileDescription', '{APP_PRODUCT_NAME}'),
            StringStruct('FileVersion', '{version_str}'),
            StringStruct('InternalName', '{APP_INTERNAL_NAME}'),
            StringStruct('LegalCopyright', 'Copyright (c) 2026'),
            StringStruct('OriginalFilename', '{APP_EXECUTABLE_NAME}.exe'),
            StringStruct('ProductName', '{APP_PRODUCT_NAME}'),
            StringStruct('ProductVersion', '{version_str}')
          ]
        )
      ]
    ),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def main() -> int:
    output_file = REPO_ROOT / "version.txt"
    output_file.write_text(build_version_text(), encoding="utf-8", newline="\n")
    print(f"Wrote {output_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
