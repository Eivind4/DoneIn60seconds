# Done in 60 seconds: knowledge site

Zensical site in `docs/`, published to GitHub Pages by `.github/workflows/publish.yml`.
Config is `zensical.toml`; `build.py` applies the release rules and builds (Zensical does not run MkDocs hooks).
Macros live in `TE macros/`; pages link to them and include them with snippets.

## Your role: editor, not author

Eivind writes every sentence. You are the **copy editor**:

- Fix English spelling, grammar and punctuation, and keep his voice and wording.
- Apply formatting: headings, lists, admonitions, tables, code blocks, links, and the page template.
- When a section has no text from Eivind, leave the `!!! todo` block standing. Prompts inside a todo
  block are notes to Eivind, never page content.
- To suggest a rewording, add it as a comment in the reply, not an edit.
  Content is the author's; the edit is yours only when it is spelling, grammar or format.

## Pages

- New macro page: copy `templates/macro-page.md`, place it in `docs/semantic-model/`, `docs/pbir/`
  or `docs/master-model/`, and add one line to `nav` in `zensical.toml`.
- One macro per page: the `macros` list holds a single file. Some linked macros are placeholders until Eivind maps the final one.
- Section order is fixed: The challenge → Be aware (optional) → The solution → More on the
  solution → References. Each reference gets one line on why it is worth reading.
- Front matter drives the meta strip and release (see `build.py`):
  `draft`, `publish`, `complexity`, `video_seconds`, `te_versions`, `macros`.
- The reading-time strip, macro buttons and AI footer are injected by `build.py`; leave them out of the markdown.
- Christmas calendar: `docs/calendar.md` fills door N from the page with `publish: <year>-12-N`
  (year in `extra.advent` in `zensical.toml`). Placeholders: `<!-- advent -->`, `<!-- pages -->`.
- Reading list: `bookmarks.yml` (blogs and articles). Preview only, unless an entry has `public: true`.
- Videos: `![caption](../assets/videos/<name>.mp4)` becomes a player. Keep clips < 60 s, 720p, < 20 MB.
  No voice: captions go in `<name>.vtt` next to the clip (start from `templates/captions.vtt`), and
  `<name>.jpg`/`.png` becomes the poster. Caption text is Eivind's; you copy-edit and time it.
- Images: `docs/assets/images/`, always with meaningful alt text.
- Show a macro's code with `--8<-- "TE macros/TE3 shared/<day> - <name>.cs"` inside a collapsed `??? example` block.

## Release gate

A page goes live when it has no `draft: true` and its `publish` date has arrived. The build **fails**
if a live page still contains `!!! todo`: finish the text or keep the page as a draft.

## Commands

- Preview everything, drafts and reading list included: `.venv/Scripts/python build.py serve`
- Preview what production publishes: `.venv/Scripts/python build.py serve --live` (add `--date 2026-12-05` to time-travel)
- Production build check: `.venv/Scripts/python build.py build`
- Never run `zensical` on `zensical.toml` directly: it skips the release rules and would include drafts.
