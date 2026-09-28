# Plan: Done in 60 seconds knowledge site

## Recommendation

| Question | Choice | Why |
|---|---|---|
| Pages, Wiki or HTML? | **GitHub Pages + MkDocs Material** | You write Markdown and get navigation, search, dark mode, warning boxes, video and code blocks. Wiki has no drafts or release control; hand-written HTML is slow to maintain. |
| Separate repo for drafts? | **No, one repo with front-matter flags** | `draft: true` hides a page; `publish: YYYY-MM-DD` releases it on that date. No copying between repos. |
| Move the macros? | **Yes, into `TE macros/` in this repo** | Pages link to them and embed the code from the same file, so there is one source of truth. |
| Videos | **MP4 in `docs/assets/videos/`** | No ads, no tracking, plays inline. Keep clips < 60 s, 720p, < 20 MB. |

## How the daily release works

1. Write pages ahead of time, each with `draft: true`.
2. When a page is finished, replace `draft: true` with `publish: 2026-10-05` (one per day).
3. Push to `main`. The workflow runs on every push and daily at 05:00 UTC, and publishes whatever is due.
4. Section pages list only released pages, so there are no broken links to unreleased ones.
5. Safety net: the build fails if a released page still has a `!!! todo` block.

## Reviewer access

- **Recommended:** add the reviewer as a collaborator. They run `SHOW_DRAFTS=1 mkdocs serve` to see
  every page as it will look, and give feedback through PR comments or issues.
  Drafts are marked with a "Draft, not published" banner.
- Note: the repo is public, so draft Markdown is readable to anyone browsing the source,
  just not on the site. If that matters, use a private staging repo instead (at the cost of copying).

## One-time setup (Eivind)

1. Repo → Settings → Pages → Source: **GitHub Actions**.
2. Review what is in `TE macros/` (two files named after a client/another person were left out).
3. Commit and push this branch (nothing is live until a page has no `draft` flag).
4. Optional: update `README.md` to link the site: https://eivind4.github.io/DoneIn60seconds/

## Writing workflow

1. Sketch: copy `templates/macro-page.md`, fill the front matter, add a nav line in `mkdocs.yml`.
2. Write your own text into each section in rough English; replace the `!!! todo` blocks.
3. Ask Claude: "copy-edit docs/…/page.md". It fixes spelling, grammar and formatting only (see `CLAUDE.md`).
4. Record the clip, save it to `docs/assets/videos/`, and preview it locally.
5. Set `publish:` and push.

## ToDo

- [ ] Video: how to add and save a script as a macro (Tabular Editor page)
- [ ] Attach the Tabular Editor 2 skill and link it (Tabular Editor page)
- [ ] Tabular Editor 2 vs 3: accept or rewrite the suggested comparison table
- [ ] PBIR files: general info page
- [ ] Semantic Link Labs: general info page
- [ ] Best Practice Rules: automate fixes with fix scripts (Tabular Editor + Semantic Link Labs)
- [ ] Logo: pick a sketch (`docs/logo-sketches.md`, preview at /logo-sketches/) and refine
- [ ] Map one final macro per page as each page is written

## Content outline

Each page is sketched as a draft with its macro linked. **Check the mapping marked ?**

**General:** About · Tabular Editor · PBIR files · Theme file · Semantic Link Labs

**Semantic model**

| Page | Macro |
|---|---|
| Create a calendar | 1 - Create English Calendar (combined), 1 - Lage Norsk kalender (combined) |
| Create time intelligence | 5 - Generate Time Intelligence DaxUDF, 3 - Time Intelligence - Calculation group ? |
| Measures with time intelligence | 5 - Create measure from model dependent functions ? |
| Dynamic axis | 5 - Create Visual calculation axis DaxUDF ? |
| Comparison | 5 - Compare first selected measure with second selected measure |
| Ranking | 5 - Create Ranking DaxUDF |
| Error bars | 5 - Create Error bars DaxUDF |
| Conditional formatting | 5 - Create Conditional formatting DaxUDF |

**PBIR**

| Page | Macro |
|---|---|
| Update column or measure name | 15 - Batch replace PBIR table column and measure references ? |
| Batch rename measures | 11 - Rename measures (TOM + PBIR sync) |
| Fix broken measures | 14 - Detect and repair broken PBIR references |
| Set display name | 12 - Set display names for selected fields in PBIR |
| Update display name | 13 - Batch rename existing display names in PBIR |
| Detect unused objects | 16 - Classify model dependencies from PBIR ? |

**Best Practice Rules:** Automate fixes with fix scripts (Tabular Editor and Semantic Link Labs)

**Master model**

| Page | Macro |
|---|---|
| Create model from perspective | 18 - Master model - create or update perspective models |

**Macros without a page (yet):** Disconnected calendar, Months to show, TopN, Pareto,
Filter historical dates, Apply conditional formatting to visuals (PBIR), Replace field in visuals (PBIR).

## Page shape

Meta strip (reading time · complexity · video length · TE version, auto) → macro button(s) → video →
The challenge → Be aware (optional) → The solution (+ collapsible script) → More on the solution
(alternatives) → References → AI note (auto).
