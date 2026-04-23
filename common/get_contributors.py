"""Fetch GitHub contributors and write them to local text files.

This script is intentionally dependency-free (stdlib only) so installer flows
can run it without extra pip packages.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess  # nosec B404
import sys
from http.client import HTTPSConnection
from pathlib import Path
from urllib.parse import urlparse


def _repo_root() -> Path:
    """Repo root.
    
    Args:
        None.
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    return Path(__file__).resolve().parents[1]


def _resolve_token() -> str | None:
    """Resolve a GitHub token from env vars or the gh CLI."""
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    gh_path = shutil.which("gh")
    if not gh_path:
        return None
    try:
        # Trusted local executable path from shutil.which; static args; no shell.
        result = subprocess.run(
            [gh_path, "auth", "token"],
            capture_output=True,
            text=True,
            check=True,
        )  # nosec B603
        return result.stdout.strip() or None
    except (OSError, subprocess.CalledProcessError):
        return None


def _github_api_request(path: str) -> tuple[int, dict[str, str], object | None]:
    """GitHub API call over explicit HTTPS, authenticated when a token is available."""
    token = _resolve_token()
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "focus-contributors-fetcher",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    conn = HTTPSConnection("api.github.com", timeout=30)
    try:
        conn.request(
            "GET",
            f"/{path.lstrip('/')}",
            headers=headers,
        )
        response = conn.getresponse()
        body = response.read().decode("utf-8")
        headers = {k.lower(): v for k, v in response.getheaders()}
        if response.status != 200:
            return response.status, headers, None
        return response.status, headers, json.loads(body)
    except (OSError, TimeoutError, json.JSONDecodeError):
        return 0, {}, None
    finally:
        conn.close()



def _github_graphql_request(query: str) -> dict | None:
    """POST a GraphQL query to the GitHub API."""
    token = _resolve_token()
    if not token:
        return None
    body = json.dumps({"query": query}).encode("utf-8")
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "focus-contributors-fetcher",
        "Authorization": f"Bearer {token}",
    }
    conn = HTTPSConnection("api.github.com", timeout=30)
    try:
        conn.request("POST", "/graphql", body=body, headers=headers)
        response = conn.getresponse()
        data = response.read().decode("utf-8")
        if response.status != 200:
            return None
        return json.loads(data)
    except (OSError, TimeoutError, json.JSONDecodeError):
        return None
    finally:
        conn.close()


def _next_page_path(link_header: str) -> str | None:
    """Extract next-page API path from RFC5988 Link header."""
    if not link_header:
        return None
    for part in link_header.split(","):
        item = part.strip()
        if 'rel="next"' not in item:
            continue
        start = item.find("<")
        end = item.find(">", start + 1)
        if start == -1 or end == -1:
            continue
        absolute = item[start + 1 : end]
        parsed = urlparse(absolute)
        if parsed.path:
            suffix = f"?{parsed.query}" if parsed.query else ""
            return f"{parsed.path.lstrip('/')}{suffix}"
    return None


def _paginated_list(path: str, max_pages: int = 20) -> list[dict]:
    """Fetch paginated GitHub list endpoints."""
    current = path
    rows: list[dict] = []
    pages = 0
    while current and pages < max_pages:
        status, headers, data = _github_api_request(current)
        if status != 200 or not isinstance(data, list):
            break
        for row in data:
            if isinstance(row, dict):
                rows.append(row)
        current = _next_page_path(headers.get("link", ""))
        pages += 1
    return rows


def _from_stats(owner: str, repo: str) -> set[str]:
    """From stats.
    
    Args:
        owner: Parameter value (str).
        repo: Parameter value (str).
    
    Returns:
        set[str]: Return value.
    
    Raises:
        None.
    """
    path = f"repos/{owner}/{repo}/stats/contributors"
    status, _, data = _github_api_request(path)
    if status == 202:
        return set()
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
    return names


def _from_contributors(owner: str, repo: str) -> set[str]:
    """From contributors.
    
    Args:
        owner: Parameter value (str).
        repo: Parameter value (str).
    
    Returns:
        set[str]: Return value.
    
    Raises:
        None.
    """
    path = f"repos/{owner}/{repo}/contributors?per_page=100&anon=1"
    data = _paginated_list(path)
    if not data:
        return set()
    names: set[str] = set()
    for row in data:
        login = row.get("login")
        if isinstance(login, str) and login:
            names.add(login.strip())
    return names


def _from_graphql_authors(owner: str, repo: str) -> set[str]:
    """Collect all commit authors — including Co-authored-by — via GraphQL.

    The GraphQL `authors` field resolves co-author email trailers to GitHub
    logins, exactly matching the Contributors graph shown on the web UI.
    """
    names: set[str] = set()
    cursor: str | None = None
    max_pages = 20
    for _ in range(max_pages):
        after = f'"{cursor}"' if cursor else "null"
        query = f"""
        {{
          repository(owner: "{owner}", name: "{repo}") {{
            defaultBranchRef {{
              target {{
                ... on Commit {{
                  history(first: 100, after: {after}) {{
                    pageInfo {{ hasNextPage endCursor }}
                    nodes {{
                      authors(first: 10) {{
                        nodes {{ user {{ login }} }}
                      }}
                    }}
                  }}
                }}
              }}
            }}
          }}
        }}
        """
        result = _github_graphql_request(query)
        if not isinstance(result, dict):
            break
        try:
            history = (
                result["data"]["repository"]["defaultBranchRef"]["target"]["history"]
            )
            for node in history.get("nodes", []):
                for author in node.get("authors", {}).get("nodes", []):
                    user = author.get("user")
                    if isinstance(user, dict):
                        login = user.get("login")
                        if isinstance(login, str) and login:
                            names.add(login.strip())
            page_info = history.get("pageInfo", {})
            if not page_info.get("hasNextPage"):
                break
            cursor = page_info.get("endCursor")
        except (KeyError, TypeError):
            break
    return names


def _write_lines(path: Path, values: list[str]) -> None:
    """Write lines.
    
    Args:
        path: Parameter value (Path).
        values: Parameter value (list[str]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    text = "\n".join(values)
    if text:
        text += "\n"
    path.write_text(text, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    """Parse args.
    
    Args:
        None.
    
    Returns:
        argparse.Namespace: Return value.
    
    Raises:
        None.
    """
    root = _repo_root()
    parser = argparse.ArgumentParser(description="Fetch and persist GitHub contributors.")
    parser.add_argument("--owner", default="ece-kalasalingam")
    parser.add_argument("--repo", default="cotas")
    parser.add_argument(
        "--output",
        default=str(root / "assets" / "about_contributors.txt"),
        help="Primary output file path.",
    )
    return parser.parse_args()


def main() -> int:
    """Main.
    
    Args:
        None.
    
    Returns:
        int: Return value.
    
    Raises:
        None.
    """
    args = parse_args()
    owner = str(args.owner).strip()
    repo = str(args.repo).strip()
    output_path = Path(str(args.output)).resolve()

    stats_names = _from_stats(owner, repo)
    contributors_names = _from_contributors(owner, repo)
    graphql_names = _from_graphql_authors(owner, repo)
    names = sorted(stats_names | contributors_names | graphql_names, key=str.lower)
    if not names:
        print(f"No contributors found for {owner}/{repo}.")
        return 1

    _write_lines(output_path, names)

    print(f"Contributors ({len(names)}): {', '.join(names)}")
    print(f"Wrote: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
