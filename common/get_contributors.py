"""Fetch GitHub contributors and write them to local text files.

This script is intentionally dependency-free (stdlib only) so installer flows
can run it without extra pip packages.
"""

from __future__ import annotations

import argparse
import json
import sys
from http.client import HTTPSConnection
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _github_api_json(path: str) -> object | None:
    """Unauthenticated GitHub API call over explicit HTTPS."""
    conn = HTTPSConnection("api.github.com", timeout=30)
    try:
        conn.request(
            "GET",
            f"/{path.lstrip('/')}",
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "focus-contributors-fetcher",
            },
        )
        response = conn.getresponse()
        if response.status != 200:
            return None
        return json.loads(response.read().decode("utf-8"))
    except (OSError, TimeoutError, json.JSONDecodeError):
        return None
    finally:
        conn.close()


def _from_stats(owner: str, repo: str) -> set[str]:
    path = f"repos/{owner}/{repo}/stats/contributors"
    data = _github_api_json(path)
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
    data = _github_api_json(path)
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
