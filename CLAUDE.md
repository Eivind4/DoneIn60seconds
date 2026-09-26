# Done in 60 seconds: knowledge site

MkDocs Material site in `docs/`, published to GitHub Pages by `.github/workflows/publish.yml`.
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
  or `docs/master-model/`, and add one line to `nav` in `mkdocs.yml`.
- One macro per page: the `macros` list holds a single file. Some linked macros are placeholders until Eivind maps the final one.
- Section order is fixed: The challenge → Be aware (optional) → The solution → More on the
  solution → References. Each reference gets one line on why it is worth reading.
- Front matter drives the meta strip and release (see `hooks/publishing.py`):
  `draft`, `publish`, `complexity`, `video_seconds`, `te_versions`, `macros`.
- The reading-time strip, macro buttons and AI footer are injected by the hook; leave them out of the markdown.
- Videos: `![caption](../assets/videos/<name>.mp4)` becomes a player. Keep clips < 60 s, 720p, < 20 MB.
- Images: `docs/assets/images/`, always with meaningful alt text.
- Show a macro's code with `--8<-- "TE macros/TE3/<file>.cs"` inside a collapsed `??? example` block.

## Release gate

A page goes live when it has no `draft: true` and its `publish` date has arrived. The build **fails**
if a live page still contains `!!! todo`: finish the text or keep the page as a draft.

## Commands

- Preview everything, drafts included: `SHOW_DRAFTS=1 .venv/Scripts/mkdocs serve`
- Production build check: `.venv/Scripts/mkdocs build`
