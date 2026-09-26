"""Publishing rules for Done in 60 seconds.

Front matter on a page decides whether it is live:
  draft: true          -> never built (unless SHOW_DRAFTS=1)
  publish: 2026-10-05  -> built from that date on (daily scheduled build releases it)

Also adds the meta strip (reading time, complexity, video length, TE versions),
macro download buttons, the TODO guard, mp4 -> <video>, and the AI footer.
"""

import datetime as dt
import math
import os
import re
from pathlib import Path
from urllib.parse import quote

import yaml
from mkdocs.exceptions import PluginError

SHOW_DRAFTS = os.environ.get("SHOW_DRAFTS") == "1"
TODAY = dt.date.today()
WORDS_PER_MINUTE = 200
FRONT_MATTER = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.S)

AI_FOOTER = (
    "\n\n---\n\n<p class=\"ai-note\">Written by Eivind Haugen. "
    "AI was used for formatting and spell checking only, not for writing the content.</p>\n"
)

_hidden: set[str] = set()  # docs-relative paths, forward slashes


def _meta(path: Path) -> dict:
    match = FRONT_MATTER.match(path.read_text(encoding="utf-8"))
    return (yaml.safe_load(match.group(1)) or {}) if match else {}


def _is_hidden(meta: dict) -> bool:
    if meta.get("draft"):
        return True
    publish = meta.get("publish")
    if isinstance(publish, str):
        publish = dt.date.fromisoformat(publish)
    return isinstance(publish, dt.date) and publish > TODAY


def _prune_nav(items):
    kept = []
    for item in items:
        if isinstance(item, dict):
            (title, value), = item.items()
            if isinstance(value, list):
                value = _prune_nav(value)
                if value:
                    kept.append({title: value})
            elif value not in _hidden:
                kept.append(item)
        elif item not in _hidden:
            kept.append(item)
    return kept


def on_config(config):
    _hidden.clear()
    docs = Path(config["docs_dir"])
    for path in docs.rglob("*.md"):
        if _is_hidden(_meta(path)):
            _hidden.add(path.relative_to(docs).as_posix())
    if not SHOW_DRAFTS and config.get("nav"):
        config["nav"] = _prune_nav(config["nav"])
    return config


def on_files(files, config):
    if SHOW_DRAFTS:
        return files
    for file in list(files):
        if file.src_uri in _hidden:
            files.remove(file)
    return files


def _reading_minutes(markdown: str) -> int:
    prose = re.sub(r"```.*?```|--8<--.*", "", markdown, flags=re.S)
    return max(1, math.ceil(len(re.findall(r"\w+", prose)) / WORDS_PER_MINUTE))


def _meta_strip(markdown: str, meta: dict) -> str:
    parts = [f":material-clock-outline: {_reading_minutes(markdown)} min read"]
    if meta.get("complexity"):
        parts.append(f":material-signal-cellular-2: {meta['complexity']}")
    if meta.get("video_seconds"):
        parts.append(f":material-play-circle-outline: {meta['video_seconds']} s video")
    if meta.get("te_versions"):
        parts.append(":material-tools: " + " · ".join(meta["te_versions"]))
    return '<div class="meta-strip" markdown>' + " &nbsp;|&nbsp; ".join(parts) + "</div>\n"


def _macro_buttons(config, meta: dict) -> str:
    repo = config.get("repo_url", "").rstrip("/")
    buttons = []
    for macro in meta.get("macros", []):
        url = f"{repo}/blob/main/{quote(macro)}"
        name = Path(macro).stem
        buttons.append(f"[:material-github: {name}]({url}){{ .md-button }}")
    return "\n".join(buttons) + "\n" if buttons else ""


def _page_list(page, config) -> str:
    """Markdown list of live sibling pages, newest release first."""
    folder = Path(config["docs_dir"], page.file.src_uri).parent
    rows = []
    for path in folder.glob("*.md"):
        uri = path.relative_to(config["docs_dir"]).as_posix()
        if path.name == "index.md" or (uri in _hidden and not SHOW_DRAFTS):
            continue
        meta = _meta(path)
        title = next(
            (l[2:].strip() for l in path.read_text(encoding="utf-8").splitlines() if l.startswith("# ")),
            path.stem,
        )
        extra = f" · {meta['complexity']}" if meta.get("complexity") else ""
        rows.append((str(meta.get("publish", "")), f"- [{title}]({path.name}){extra}"))
    rows.sort(reverse=True)
    return "\n".join(r for _, r in rows) or "*First page coming soon.*"


def on_page_markdown(markdown, page, config, files):
    meta = page.meta
    live = page.file.src_uri not in _hidden
    markdown = markdown.replace("<!-- pages -->", _page_list(page, config))

    if live and not SHOW_DRAFTS and re.search(r"^!!! todo", markdown, re.M):
        raise PluginError(
            f"{page.file.src_uri} is scheduled/published but still has '!!! todo' blocks."
        )

    header = ""
    if not live:
        header += '!!! warning "Draft, not published"\n    Only visible in preview.\n\n'
    if meta.get("complexity") or meta.get("video_seconds"):
        header += _meta_strip(markdown, meta)
    header += _macro_buttons(config, meta)

    if header:
        lines = markdown.split("\n")
        h1 = next((i for i, line in enumerate(lines) if line.startswith("# ")), -1)
        lines.insert(h1 + 1, "\n" + header)
        markdown = "\n".join(lines)

    return markdown + AI_FOOTER


def on_page_content(html, page, config, files):
    # ![caption](clip.mp4) -> inline video player
    return re.sub(
        r'<img alt="([^"]*)" src="([^"]+\.mp4)"\s*/?>',
        r'<video controls muted playsinline preload="metadata" title="\1">'
        r'<source src="\2" type="video/mp4"></video>',
        html,
    )
