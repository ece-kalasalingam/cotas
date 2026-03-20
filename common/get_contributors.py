"""Fetch GitHub contributors and write them to local text files.

This script is intentionally dependency-free (stdlib only) so installer flows
can run it without extra pip packages.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _gh_api_json(path: str) -> object | None:
    """Try GitHub CLI first (handles auth/session better)."""
    cmd = ["gh", "api", path]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
    except Exception:
        return None
    if proc.returncode != 0:
        return None
    try:
        return json.loads(proc.stdout)
    except Exception:
        return None


def _http_json(url: str) -> object | None:
    """Fallback unauthenticated HTTP call."""
    req = urllib.request.Request(
        url=url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "focus-contributors-fetcher",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return None


def _from_stats(owner: str, repo: str) -> set[str]:
    path = f"repos/{owner}/{repo}/stats/contributors"
    data = _gh_api_json(path)
    if data is None:
        data = _http_json(f"https://api.github.com/{path}")
    if not isinstance(data, list):
        return set()
    names: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        author = row.get("author")
        if isinstance(author, dict):
            login = author.get("login")
            if isinstance(login, str) and login:
                names.add(login.strip())
                continue
        author_name = row.get("author_name")
        if isinstance(author_name, str) and author_name:
            names.add(author_name.strip())
    return names


def _from_contributors(owner: str, repo: str) -> set[str]:
    path = f"repos/{owner}/{repo}/contributors?per_page=100"
    data = _gh_api_json(path)
    if data is None:
        data = _http_json(f"https://api.github.com/{path}")
    if not isinstance(data, list):
        return set()
    names: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        login = row.get("login")
        if isinstance(login, str) and login:
            names.add(login.strip())
    return names


def _write_lines(path: Path, values: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = "\n".join(values)
    if text:
        text += "\n"
    path.write_text(text, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    root = _repo_root()
    parser = argparse.ArgumentParser(description="Fetch and persist GitHub contributors.")
    parser.add_argument("--owner", default="ece-kalasalingam")
    parser.add_argument("--repo", default="cotas")
    parser.add_argument(
        "--output",
        default=str(root / "assets" / "about_contributors.txt"),
        help="Primary output file path.",
    )
    parser.add_argument(
        "--legacy-output",
        default=str(root / "common" / "contributors_list.txt"),
        help="Optional secondary output file path.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    owner = str(args.owner).strip()
    repo = str(args.repo).strip()
    output_path = Path(str(args.output)).resolve()
    legacy_output = Path(str(args.legacy_output)).resolve()

    stats_names = _from_stats(owner, repo)
    contributors_names = _from_contributors(owner, repo)
    names = sorted(stats_names | contributors_names, key=str.lower)
    if not names:
        print(f"No contributors found for {owner}/{repo}.")
        return 1

    _write_lines(output_path, names)
    if legacy_output != output_path:
        _write_lines(legacy_output, names)

    print(f"Contributors ({len(names)}): {', '.join(names)}")
    print(f"Wrote: {output_path}")
    if legacy_output != output_path:
        print(f"Wrote: {legacy_output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
