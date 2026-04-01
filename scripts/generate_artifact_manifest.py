"""Generate SHA256 manifest for release artifacts."""

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
    parser = argparse.ArgumentParser(description="Generate artifact manifest (SHA256).")
    parser.add_argument("artifacts", nargs="+", help="Artifact file paths.")
    parser.add_argument("--out", default="artifact-manifest.json", help="Output manifest file.")
    args = parser.parse_args()

    entries: list[dict[str, object]] = []
    for raw_path in args.artifacts:
        path = Path(raw_path)
        if not path.exists() or not path.is_file():
            raise SystemExit(f"Artifact not found: {path}")
        entries.append(
            {
                "path": str(path.as_posix()),
                "size": path.stat().st_size,
                "sha256": _sha256(path),
            }
        )

    output = {
        "schema_version": 1,
        "artifacts": entries,
    }
    Path(args.out).write_text(json.dumps(output, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

