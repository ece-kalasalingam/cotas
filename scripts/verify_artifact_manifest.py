"""Verify release artifacts against a SHA256 manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def _sha256(path: Path) -> str:
    """Sha256.
    
    Args:
        path: Parameter value (Path).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    """Main.
    
    Args:
        None.
    
    Returns:
        int: Return value.
    
    Raises:
        None.
    """
    parser = argparse.ArgumentParser(description="Verify artifact manifest.")
    parser.add_argument("manifest", help="Manifest JSON path.")
    args = parser.parse_args()

    manifest_path = Path(args.manifest)
    if not manifest_path.exists():
        raise SystemExit(f"Manifest not found: {manifest_path}")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    artifacts = payload.get("artifacts", [])
    if not isinstance(artifacts, list):
        raise SystemExit("Invalid manifest format: artifacts must be a list.")

    for item in artifacts:
        if not isinstance(item, dict):
            raise SystemExit("Invalid manifest entry.")
        path = Path(str(item.get("path", "")))
        expected_hash = str(item.get("sha256", ""))
        if not path.exists():
            raise SystemExit(f"Missing artifact: {path}")
        actual_hash = _sha256(path)
        if actual_hash != expected_hash:
            raise SystemExit(f"Checksum mismatch for {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

