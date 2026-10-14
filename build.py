"""Build and preview Done in 60 seconds with Zensical.

    python build.py serve           preview at http://localhost:8000, drafts and reading list included
    python build.py serve --live    preview exactly what production will publish
    python build.py build           production build into .build/site (used by the workflow)
    add --date 2026-12-05           to see what is released on that date

Zensical does not run MkDocs hooks, so build.py applies the release rules before it builds.
It copies docs/ to .build/docs, transforms each page and writes .build/mkdocs.yml:

  draft: true          -> left out (shown with a banner in preview)
  publish: 2026-12-01  -> left out until that date (the daily scheduled build releases it)

It also adds the meta strip, macro buttons, video players, the AI footer,
<!-- pages --> lists, the <!-- advent --> calendar and the reading list,
and fails the build when a live page still contains '!!! todo'.
"""

import argparse
import datetime as dt
import math
import os
import re
import shutil
import subprocess
import sys
import time
import tomllib
from pathlib import Path
from urllib.parse import quote

import yaml

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
BUILD = ROOT / ".build"
OUT_DOCS = BUILD / "docs"
SOURCE_CONFIG = ROOT / "zensical.toml"
BUILD_CONFIG = BUILD / "mkdocs.yml"
BOOKMARKS = ROOT / "bookmarks.yml"
ZENSICAL = [sys.executable, "-m", "zensical"]

TODAY = dt.date.today()
WORDS_PER_MINUTE = 200
FRONT_MATTER = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.S)
VIDEO = re.compile(r"!\[([^\]]*)\]\(([^)\s]+\.mp4)\)")

AI_FOOTER = (
    "\n\n---\n\n<p class=\"ai-note\">Written by Eivind Haugen. "
    "AI was used for formatting and spell checking only, not for writing the content.</p>\n"
)


class ReleaseError(Exception):
    pass


# ---------------------------------------------------------------------------
# Front matter and release rules
# ---------------------------------------------------------------------------


def _split(text: str) -> tuple[dict, str]:
    match = FRONT_MATTER.match(text)
    if not match:
        return {}, text
    return yaml.safe_load(match.group(1)) or {}, text[match.end():]


def _publish_date(meta: dict) -> dt.date | None:
    publish = meta.get("publish")
    if isinstance(publish, str):
        publish = dt.date.fromisoformat(publish)
    return publish if isinstance(publish, dt.date) else None


def _is_hidden(meta: dict) -> bool:
    publish = _publish_date(meta)
    return bool(meta.get("draft")) or (publish is not None and publish > TODAY)


def _title(body: str, fallback: str) -> str:
    return next((l[2:].strip() for l in body.splitlines() if l.startswith("# ")), fallback)


class Page:
    def __init__(self, path: Path):
        self.path = path
        self.uri = path.relative_to(DOCS).as_posix()
        self.meta, self.body = _split(path.read_text(encoding="utf-8"))
        self.title = _title(self.body, path.stem)
        self.hidden = _is_hidden(self.meta)
        self.publish = _publish_date(self.meta)


def _rel(target: str, source: str) -> str:
    """Relative Markdown link from page `source` to page `target` (docs-relative URIs)."""
    return os.path.relpath(target, os.path.dirname(source) or ".").replace(os.sep, "/")


# ---------------------------------------------------------------------------
# Page transformations
# ---------------------------------------------------------------------------


def _reading_minutes(markdown: str) -> int:
    prose = re.sub(r"```.*?```|--8<--.*", "", markdown, flags=re.S)
    return max(1, math.ceil(len(re.findall(r"\w+", prose)) / WORDS_PER_MINUTE))


def _meta_strip(page: Page) -> str:
    meta = page.meta
    parts = [f":lucide-clock: {_reading_minutes(page.body)} min read"]
    if meta.get("complexity"):
        parts.append(f":lucide-gauge: {meta['complexity']}")
    if meta.get("video_seconds"):
        parts.append(f":lucide-circle-play: {meta['video_seconds']} s video")
    if meta.get("te_versions"):
        parts.append(":lucide-wrench: " + " · ".join(meta["te_versions"]))
    items = "".join(f"<span>{p}</span>" for p in parts)
    return f'<div class="meta-strip" markdown>{items}</div>\n'


def _macro_buttons(config: dict, meta: dict) -> str:
    repo = (config.get("repo_url") or "").rstrip("/")
    buttons = []
    for macro in meta.get("macros") or []:
        url = f"{repo}/blob/main/{quote(macro)}"
        buttons.append(f"[:fontawesome-brands-github: {Path(macro).stem}]({url}){{ .md-button }}")
    return "\n".join(buttons) + "\n" if buttons else ""


def _videos(markdown: str) -> str:
    # ![caption](clip.mp4) -> inline player. Zensical rewrites the relative src for directory URLs.
    def player(match: re.Match) -> str:
        caption, src = match.groups()
        return (
            f'<figure class="video"><video controls muted playsinline preload="metadata" '
            f'title="{caption}"><source src="{src}" type="video/mp4"></video></figure>'
        )

    return VIDEO.sub(player, markdown)


def _page_list(page: Page, pages: dict[str, Page], preview: bool) -> str:
    """Markdown list of live sibling pages in calendar order (publish date); undated pages last."""
    folder = os.path.dirname(page.uri)
    rows = []
    for other in pages.values():
        if os.path.dirname(other.uri) != folder or other.path.name == "index.md":
            continue
        if other.hidden and not preview:
            continue
        extra = f" · {other.meta['complexity']}" if other.meta.get("complexity") else ""
        rows.append((str(other.publish or "9999"), f"- [{other.title}]({other.path.name}){extra}"))
    rows.sort()
    return "\n".join(r for _, r in rows) or "*First page coming soon.*"


def _icon(page: Page) -> str:
    """Front matter `icon: lucide/gift` as an inline icon for the calendar door."""
    icon = page.meta.get("icon")
    return f'<span class="advent__icon">:{icon.replace("/", "-")}:</span>' if icon else ""


def _advent(page: Page, pages: dict[str, Page], config: dict, preview: bool) -> str:
    """Grid of calendar doors. Door N opens on N December when a live page is published that day."""
    advent = config.get("extra", {}).get("advent", {})
    year, days = advent.get("year", TODAY.year), advent.get("days", 24)
    by_date = {p.publish: p for p in pages.values() if p.publish}
    doors = []
    for day in range(1, days + 1):
        date = dt.date(year, 12, day)
        target = by_date.get(date)
        classes = ".advent__door" + (" .is-today" if date == TODAY else "")
        today = " is-today" if date == TODAY else ""
        number = f'<span class="advent__day">{day}</span>'
        if target and not target.hidden:
            label = f'{_icon(target)}<span class="advent__title">{target.title}</span>'
            doors.append(f"[{number}{label}]({_rel(target.uri, page.uri)}){{ {classes} .is-open }}")
        elif preview:
            # Planning view: show what is behind each door, and which days are still empty
            if target:
                label = f'{_icon(target)}<span class="advent__title">{target.title}</span>'
                doors.append(
                    f"[{number}{label}]({_rel(target.uri, page.uri)}){{ {classes} .is-planned }}"
                )
            else:
                label = '<span class="advent__title">Nothing planned yet</span>'
                doors.append(f'<span class="advent__door is-empty{today}">{number}{label}</span>')
        else:
            text = "Coming soon" if date <= TODAY else f"Opens {day} December"
            label = f'<span class="advent__title">{text}</span>'
            doors.append(f'<span class="advent__door is-locked{today}">{number}{label}</span>')
    return '<div class="advent" markdown>\n\n' + "\n".join(doors) + "\n\n</div>"


def _transform(page: Page, pages: dict[str, Page], config: dict, preview: bool) -> str:
    markdown = page.body
    markdown = markdown.replace("<!-- pages -->", _page_list(page, pages, preview))
    if "<!-- advent -->" in markdown:
        markdown = markdown.replace("<!-- advent -->", _advent(page, pages, config, preview))

    if not page.hidden and not preview and re.search(r"^!!! todo", markdown, re.M):
        raise ReleaseError(f"{page.uri} is scheduled/published but still has '!!! todo' blocks.")

    header = ""
    if page.hidden:
        when = f"scheduled for {page.publish}" if page.publish and not page.meta.get("draft") else "draft"
        header += f'!!! warning "Not published ({when})"\n    Only visible in preview.\n\n'
    if page.meta.get("complexity") or page.meta.get("video_seconds"):
        header += _meta_strip(page)
    header += _macro_buttons(config, page.meta)

    if header:
        lines = markdown.split("\n")
        h1 = next((i for i, line in enumerate(lines) if line.startswith("# ")), -1)
        lines.insert(h1 + 1, "\n" + header)
        markdown = "\n".join(lines)

    markdown = _videos(markdown)
    if not page.meta.get("hide_ai_note"):
        markdown += AI_FOOTER

    front = yaml.safe_dump(page.meta, sort_keys=False, allow_unicode=True) if page.meta else ""
    return (f"---\n{front}---\n\n" if front else "") + markdown


# ---------------------------------------------------------------------------
# Reading list (bookmarks.yml)
# ---------------------------------------------------------------------------


def _link_row(entry: dict) -> str:
    name = entry.get("title") or entry.get("name") or entry["url"]
    parts = [f"[{name}]({entry['url']})"]
    if entry.get("author"):
        parts.append(f"*{entry['author']}*")
    if entry.get("feed"):
        parts.append(f"[:lucide-rss:]({entry['feed']} \"RSS feed\")")
    if entry.get("note"):
        parts.append(entry["note"])
    return "- " + " · ".join(parts)


def _by_topic(entries: list[dict]) -> str:
    topics: dict[str, list[dict]] = {}
    for entry in entries:
        for topic in entry.get("topics") or ["Other"]:
            topics.setdefault(topic, []).append(entry)
    out = []
    for topic in sorted(topics, key=str.casefold):
        rows = sorted(topics[topic], key=lambda e: str(e.get("added", "")), reverse=True)
        out.append(f"### {topic}\n\n" + "\n".join(_link_row(e) for e in rows))
    return "\n\n".join(out)


def _reading_pages(preview: bool) -> dict[str, str]:
    """docs-relative path -> Markdown. Private entries only appear in preview."""
    if not BOOKMARKS.exists():
        return {}
    data = yaml.safe_load(BOOKMARKS.read_text(encoding="utf-8")) or {}
    blogs, articles = data.get("blogs") or [], data.get("articles") or []
    pages = {}

    if preview:
        unread = [a for a in articles if not a.get("read")]
        read = [a for a in articles if a.get("read")]
        md = [
            "---\nhide_ai_note: true\n---\n\n# Reading list",
            '!!! info "Private"\n    Only visible in preview. Edit `bookmarks.yml` to add links.',
            "## Blogs to follow\n\n" + (_by_topic(blogs) or "*None yet.*"),
            "## To read\n\n" + (_by_topic(unread) or "*None yet.*"),
        ]
        if read:
            md.append("## Read\n\n" + _by_topic(read))
        pages["reading-list.md"] = "\n\n".join(md) + "\n"

    shared = [b for b in blogs if b.get("public")] + [a for a in articles if a.get("public")]
    if shared:
        md = ["---\nhide_ai_note: true\n---\n\n# Worth following", _by_topic(shared)]
        pages["worth-following.md"] = "\n\n".join(md) + "\n"
    return pages


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------


def _prune_nav(items, keep: set[str]):
    kept = []
    for item in items:
        if isinstance(item, dict):
            (title, value), = item.items()
            if isinstance(value, list):
                value = _prune_nav(value, keep)
                if value:
                    kept.append({title: value})
            elif value in keep or "://" in value:
                kept.append(item)
        elif item in keep:
            kept.append(item)
    return kept


def _config(source: dict, keep: set[str], extra_pages: dict[str, str]) -> dict:
    config = dict(source)
    config["docs_dir"] = "docs"
    config["site_dir"] = "site"
    theme = dict(config.get("theme", {}))
    theme["custom_dir"] = "../overrides"
    config["theme"] = theme

    nav = list(config.get("nav", []))
    for path, title in (("worth-following.md", "Worth following"), ("reading-list.md", "Reading list")):
        if path in extra_pages:
            nav.append({title: path})
    config["nav"] = _prune_nav(nav, keep)

    md = dict(config.get("markdown_extensions", {}))
    snippets = dict(md.get("pymdownx.snippets", {}))
    snippets["base_path"] = [ROOT.as_posix()]
    md["pymdownx.snippets"] = snippets
    config["markdown_extensions"] = md
    return config


def _flatten_extensions(md: dict) -> dict:
    """TOML dotted keys load as nested tables (pymdownx -> {emoji -> {...}}); flatten for mkdocs.yml."""
    flat = {}
    for name, value in md.items():
        if name == "pymdownx" and isinstance(value, dict):
            for sub, options in value.items():
                flat[f"pymdownx.{sub}"] = options
        else:
            flat[name] = value
    return flat


# ---------------------------------------------------------------------------
# Prepare
# ---------------------------------------------------------------------------


def _write(path: Path, content: str) -> None:
    """Write only when changed, so the preview server rebuilds just what changed."""
    if path.exists() and path.read_text(encoding="utf-8") == content:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def prepare(preview: bool) -> None:
    with SOURCE_CONFIG.open("rb") as f:
        source = tomllib.load(f).get("project", {})
    source["markdown_extensions"] = _flatten_extensions(source.get("markdown_extensions", {}))

    pages = {p.uri: p for p in (Page(path) for path in sorted(DOCS.rglob("*.md")))}
    extra_pages = _reading_pages(preview)

    written: set[Path] = set()
    keep: set[str] = set()
    for uri, page in pages.items():
        if page.hidden and not preview:
            continue
        target = OUT_DOCS / uri
        _write(target, _transform(page, pages, source, preview))
        written.add(target)
        keep.add(uri)
    for uri, markdown in extra_pages.items():
        target = OUT_DOCS / uri
        _write(target, markdown)
        written.add(target)
        keep.add(uri)

    for path in DOCS.rglob("*"):
        if path.is_file() and path.suffix != ".md":
            target = OUT_DOCS / path.relative_to(DOCS)
            if not target.exists() or target.stat().st_mtime < path.stat().st_mtime:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, target)
            written.add(target)

    # Remove pages that were unpublished, renamed or deleted since the last run
    if OUT_DOCS.exists():
        for path in OUT_DOCS.rglob("*"):
            if path.is_file() and path not in written:
                path.unlink()

    config = _config(source, keep, extra_pages)
    _write(BUILD_CONFIG, "# Generated by build.py from zensical.toml. Do not edit.\n"
           + yaml.safe_dump(config, sort_keys=False, allow_unicode=True))


def _snapshot() -> dict[Path, float]:
    watched = [DOCS, ROOT / "TE macros", ROOT / "overrides"]
    files = [p for d in watched if d.exists() for p in d.rglob("*") if p.is_file()]
    files += [p for p in (SOURCE_CONFIG, BOOKMARKS) if p.exists()]
    return {p: p.stat().st_mtime for p in files}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["build", "serve"])
    parser.add_argument("--live", action="store_true", help="serve: hide drafts, as in production")
    parser.add_argument("--drafts", action="store_true", help="build: include drafts (never for publishing)")
    parser.add_argument("--date", type=dt.date.fromisoformat, help="pretend today is this date (YYYY-MM-DD)")
    args = parser.parse_args()

    global TODAY
    TODAY = args.date or TODAY

    preview = args.drafts if args.command == "build" else not args.live
    try:
        prepare(preview)
    except ReleaseError as error:
        print(f"Release check failed: {error}", file=sys.stderr)
        return 1

    if args.command == "build":
        return subprocess.call([*ZENSICAL, "build", "--clean", "-f", str(BUILD_CONFIG)])

    server = subprocess.Popen([*ZENSICAL, "serve", "-f", str(BUILD_CONFIG)])
    state = _snapshot()
    try:
        while server.poll() is None:
            time.sleep(1)
            current = _snapshot()
            if current != state:
                state = current
                try:
                    prepare(preview)
                except ReleaseError as error:
                    print(f"Release check failed: {error}", file=sys.stderr)
                except Exception as error:  # keep the preview running while a page is half-written
                    print(f"Preview update failed: {error}", file=sys.stderr)
    except KeyboardInterrupt:
        server.terminate()
    return server.wait()


if __name__ == "__main__":
    sys.exit(main())
