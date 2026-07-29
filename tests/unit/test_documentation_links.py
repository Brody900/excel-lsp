from __future__ import annotations

import re
from collections import Counter
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).parents[2]
SKIP_DIRS = {
    ".git",
    ".hypothesis",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    ".venv",
    "build",
    "dist",
}
LINK_RE = re.compile(r"(?<!!)\[[^\]]+\]\(([^)]+)\)")
HEADING_RE = re.compile(r"^#{1,6}\s+(.+?)\s*#*\s*$", re.MULTILINE)
HTML_ID_RE = re.compile(r"\bid=[\"']([^\"']+)[\"']", re.IGNORECASE)
FENCE_RE = re.compile(r"^```.*?^```\s*$", re.MULTILINE | re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")


def _markdown_files() -> list[Path]:
    return sorted(
        path
        for path in ROOT.rglob("*.md")
        if not any(part in SKIP_DIRS for part in path.relative_to(ROOT).parts)
    )


def _github_slug(heading: str) -> str:
    plain = TAG_RE.sub("", heading).strip().lower()
    plain = re.sub(r"[^\w\- ]", "", plain, flags=re.UNICODE)
    return re.sub(r" +", "-", plain)


def _anchors(text: str) -> set[str]:
    anchors = set(HTML_ID_RE.findall(text))
    seen: Counter[str] = Counter()
    for heading in HEADING_RE.findall(FENCE_RE.sub("", text)):
        base = _github_slug(heading)
        suffix = seen[base]
        seen[base] += 1
        anchors.add(base if suffix == 0 else f"{base}-{suffix}")
    return anchors


def test_all_repository_markdown_links_and_anchors_resolve() -> None:
    failures: list[str] = []
    for source in _markdown_files():
        text = FENCE_RE.sub("", source.read_text(encoding="utf-8"))
        for raw_target in LINK_RE.findall(text):
            target = raw_target.strip().split(maxsplit=1)[0].strip("<>")
            if target.startswith(("http://", "https://", "mailto:")):
                continue
            path_text, separator, fragment = target.partition("#")
            destination = (
                source if not path_text else (source.parent / unquote(path_text)).resolve()
            )
            try:
                destination.relative_to(ROOT)
            except ValueError:
                failures.append(f"{source.relative_to(ROOT)}: path escapes repository: {target}")
                continue
            if not destination.exists():
                failures.append(f"{source.relative_to(ROOT)}: missing target: {target}")
                continue
            if separator and destination.suffix.lower() == ".md":
                available = _anchors(destination.read_text(encoding="utf-8"))
                if unquote(fragment).lower() not in available:
                    failures.append(f"{source.relative_to(ROOT)}: missing anchor: {target}")
    assert not failures, "\n" + "\n".join(failures)
