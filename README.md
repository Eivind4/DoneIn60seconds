# DoneIn60seconds

Solving common Power BI challenges in less than 60 seconds: deterministic, token free and model-independent.

This repository holds the Tabular Editor macros, Semantic Link Labs notebooks and session material behind the *Done in 60 seconds* community sessions, together with the knowledge site that documents them at **[eivind4.github.io/DoneIn60seconds](https://eivind4.github.io/DoneIn60seconds/)**. The content covers time intelligence, calendars, DAX User-Defined Functions (UDFs), PBIR report automation, usage analysis and cleanup, best-practice analysis, AI readiness and master models.

---

## Repository structure

```
docs/                            ← knowledge site pages (Zensical)
  general/                       ← Tabular Editor, PBIR files, theme file, Semantic Link Labs
  semantic-model/                ← calendar, time intelligence, ranking, error bars, ...
  pbir/                          ← report renames, broken references, usage dependency
  best-practice-rules/           ← BPA fix scripts, date inconsistencies, duplicates
  ai-readiness/                  ← AI settings, tailored BPA rules, AI schema
  master-model/                  ← perspective models and mini models
  calendar.md                    ← Christmas calendar (one page per December day)
TE macros/
  TE3/                           ← current macros, grouped by topic (see below)
  Conferences/
    New Stars of Data 2025/      ← session scripts (TE2 compatible)
    Data Toboggan - Summer 2025/ ← session scripts + presentation
  old/                           ← earlier TE2 and TE3 versions, kept for reference
Dax UDF/                         ← TMDL snippets for DAX UDF patterns
Semantic Link Labs/
  BPA/                           ← Fabric notebook for BPA with SLL
  BPA-SemanticLinkLabsReport/    ← PBIR report for BPA results
  nb_Copilot_model_settings_editor.py ← Copilot settings editor notebook
templates/macro-page.md          ← template for a new macro page
build.py, zensical.toml          ← site build and configuration
bookmarks.yml                    ← reading list (preview only unless marked public)
```

---

## Knowledge site

The site is built with [Zensical](https://zensical.org/) and published to GitHub Pages by `.github/workflows/publish.yml` on every push to `main` and once a day at 05:00 UTC. `build.py` applies the release rules before building: a page is left out while it has `draft: true` or a `publish` date in the future, and the build fails if a live page still contains a `!!! todo` block. The daily build releases pages on their publish date, which also fills the doors of the Christmas calendar.

Pages follow `templates/macro-page.md`, and each page documents one macro. Front matter (`draft`, `publish`, `complexity`, `video_seconds`, `te_versions`, `macros`) drives the meta strip, the macro buttons and the release date.

```
pip install -r requirements.txt
python build.py serve                  # preview with drafts and reading list
python build.py serve --live           # preview what production publishes
python build.py serve --live --date 2026-12-05   # preview a future release date
python build.py build                  # production build into .build/site
```

Do not run `zensical` on `zensical.toml` directly: it skips the release rules and would publish drafts.

---

## TE macros — TE3

Current macros for **Tabular Editor 3**, grouped by topic. Most of them use TE3 features such as DAX UDFs and Calendars, and the PBIR macros edit report files in a PBIP project on disk.

> [!WARNING]
> **The PBIR macros are experimental for now.** This covers every macro that reads or edits PBIR report files: `00 Info/PBIR`, `06 Report (PBIR)`, `07 Usage and cleanup` and `Perspective from report.cs`. Run them on a copy of the project or under source control, and check the report in Power BI Desktop afterwards.

### 00 Info

| File | Purpose |
|------|---------|
| `BPA/Export BPA results.cs` | Exports Best Practice Analyzer results (based on m-kovalsky's script) |
| `BPA/Export BPA rules.cs` | Exports the BPA rules in use |
| `BPA/Load BPA rules - Analysis Services.cs` | Loads a BPA rule set into the model |
| `PBIR/List Fabric App usage.cs` | Read-only overview of which measures and columns Fabric Apps query from the model |
| `PBIR/Organize columns by usage (PBIR).cs` | Puts columns in display folders by how they are used in reports and the model |
| `PBIR/Organize measures by usage (PBIR and Fabric Apps).cs` | Puts measures in display folders by how they are used in PBIR reports and Fabric Apps |

### 01 Tables

| File | Purpose |
|------|---------|
| `Create a formatted date table with calendar options.cs` | Creates a calendar table, marks it as a date table and configures Gregorian and ISO/weekly calendar options in one run |

### 02 DaxUDF

| File | Purpose |
|------|---------|
| `Generate Time Intelligence DaxUDF.cs` | Generates time intelligence UDFs (accumulated, moving totals and averages, semi-additive, date filtering), in date-column or Calendar mode |
| `Rank by selected columns DaxUDF.cs` | Ranking UDFs for a base measure by one or more selected columns |
| `Create Pareto Analysis DaxUDF - model independent.cs` | Model-independent UDFs for the dynamic Pareto analysis pattern |
| `Conditional formatting by thresholds DaxUDF.cs` | UDFs for background colour, font colour and good/bad/neutral flags based on thresholds |
| `Error bars for actual vs target DaxUDF.cs` | UDFs that drive error bars comparing an actual measure with a target |
| `Create Visual calculation axis DaxUDF.cs` | UDFs that set a dynamic axis min/max in visual calculations |

### 03 Calculation Groups

| File | Purpose |
|------|---------|
| `Dynamic axis from slicer selection.cs` | Creates or updates a "Period to show" calculation group built on DAX UDFs |

### 04 Measures

| File | Purpose |
|------|---------|
| `Create measure from model dependent functions.cs` | Creates measures for selected base measures from the model-dependent UDFs, using their suffix and folder annotations |
| `Compare first selected measure with second selected measure DaxUDF.cs` | Creates comparison measures between two selected measures |

### 06 Report (PBIR) (experimental)

| File | Purpose |
|------|---------|
| `Rename measures (TOM + PBIR sync).cs` | Renames measures in the model and updates every reference in the PBIR report files |
| `Batch replace PBIR table column and measure references.cs` | Replaces table, column and measure references across the model and one or more reports |
| `Detect and repair broken PBIR references.cs` | Finds report references that no longer exist in the model and repairs them |
| `Replace field in visuals (PBIR).cs` | Swaps one measure or column for another in report visuals without changing the model |
| `Set display names for selected fields in PBIR.cs` | Sets explicit display names for selected fields across reports |
| `Batch rename existing display names in PBIR.cs` | Standardizes existing display names by their underlying field |

### 07 Usage and cleanup

| File | Purpose |
|------|---------|
| `Classify model dependencies from PBIR and Fabric Apps.cs` | Read-only overview of how the model is used by reports and Fabric Apps |
| `Merge duplicate measures (PBIR and Fabric Apps).cs` | Finds measures with the same DAX, keeps one and repoints reports and Fabric Apps |

### 08 Master models

| File | Purpose |
|------|---------|
| `Master model - create or update perspective models.cs` | Builds smaller semantic models from the perspectives of one master model (PBIP) |
| `Perspective from report.cs` | Creates a perspective with exactly the fields one or more PBIR reports use |

### 09 Documentation

| File | Purpose |
|------|---------|
| `Create model documentation (Info.view).cs` | Creates a calculated documentation table from the INFO views |

---

## TE macros — Conferences

### New Stars of Data 2025 (May 16, 2025)

Introductory session on Tabular Editor automation. Scripts are compatible with the free open-source **Tabular Editor 2**.

| # | Script | Purpose |
|---|--------|---------|
| 1 | `1. Create calculated date table.csx` | Generates a calculated DAX date table |
| 1 | `1. Create calculated measure table.csx` | Creates an empty measure staging table |
| 2 | `2. Mark as date-table and create measures and relationship.csx` | Marks the date table and wires up a relationship |
| 2 | `2. Organize Calculated date-table.csx` | Organises columns in a calculated date table |
| 2 | `2. Organize Contoso date-table.csx` | Organises columns for the Contoso sample date table |
| 3 | `3. Create Time Intelligence Calculation group.csx` | Creates a time intelligence calculation group |
| 3 | `3. Create or Add to Time Intelligence Calculation group.csx` | Creates or extends an existing calculation group |
| 3 | `3. Add items to Time Intelligence Calculation group.csx` | Adds individual items to an existing group |
| 4 | `4. Autogenerate SUM for selected measures.csx` | Auto-generates SUM measures for selected columns |
| 4 | `4. Create deviation measure.csx` | Creates a YoY or period deviation measure |
| 4 | `4. Create Time Intelligence measures.csx` | Generates TI measures for selected base measures |
| 8 | `8. Create model documentation.csx` | Exports model metadata as documentation |
| 9 | `9. Do not summarize numeric columns.csx` | Sets summarization to None on numeric columns |
| 9 | `9. Format measures - Max 20.csx` | Formats up to 20 selected measures |
| 9 | `9. Format string for date-columns.csx` | Applies a consistent format string to date columns |
| 9 | `9. Load best practice rules.csx` | Imports BPA rules from a remote source |
| 9 | `9. Set Available in MDX to false.csx` | Hides selected objects from MDX clients |
| 9 | `9. Set Available in MDX to false for all.csx` | Hides all eligible objects from MDX clients |
| — | `MacroActions.json` | Macro action file — import into TE2 to register all scripts as macros |

### Data Toboggan — Summer 2025 (July 12, 2025)

Session focused on intelligent time intelligence using Calculation Groups and Calendars. See the included PDF and the YouTube recording for a walkthrough.

| File | Purpose |
|------|---------|
| `Time Intelligence Calculation group.csx` | Creates a time intelligence calculation group |
| `Create measures with Time Intelligence calc group logic.csx` | Generates measures that leverage the calc group |
| `Format Date Table csharp.csx` | Formats a date table (columns, descriptions, folders) |
| `Format all date-columns` | Bulk-formats all date columns in the model |
| `Format Calendar.tmdl` | TMDL snippet for a formatted Calendar object |
| `Time Intelligence calculation group.tmdl` | TMDL snippet for the time intelligence calc group |
| `Months to Show.tmdl` | TMDL snippet for a "Months to Show" what-if parameter |
| `Data Toboggan - Done in 60 seconds - Intelligent time intelligence.pdf` | Slide deck from the session |

---

## TE macros — old

`TE macros/old/` keeps earlier, numbered versions of the macros for both TE2 and TE3, including the TE2-compatible DAX UDF generators and the Norwegian calendar scripts. Use the macros in `TE macros/TE3/` for new work.

---

## Dax UDF

TMDL reference snippets for DAX UDF patterns used across scripts.

| File | Purpose |
|------|---------|
| `Comparison Function.tmdl` | TMDL definition of a comparison DAX UDF (difference and % difference between two measures). |

---

## Semantic Link Labs

Resources for running analysis and administration at scale with the [Semantic Link Labs](https://github.com/microsoft/semantic-link-labs) Python library inside Microsoft Fabric.

| File | Purpose |
|------|---------|
| `BPA/Fabric Februay - SLL BPA demo Shared.ipynb` | Fabric notebook demonstrated at the Fabric February community event. Runs BPA rules across one or more semantic models and writes results to a Lakehouse table. |
| `BPA-SemanticLinkLabsReport/` | Power BI Enhanced Report (PBIR) for visualising the BPA results produced by the notebook above. |
| `nb_Copilot_model_settings_editor.py` | Fabric Python notebook that turns *Copilot and MCP tools for readers* and *Approved for Copilot* on or off per semantic model, and shows whether each model is prepped for AI (AI instructions, AI schema, verified answers or a data agent). Needs write permission on the model; it uses an undocumented endpoint, so test on one model first. |

---

## Getting started

1. **Tabular Editor 3**: open any `.cs` file from `TE macros/TE3/` in the TE3 script editor and run it, or save it as a macro. The site page for each macro explains when and how to use it.
2. **Tabular Editor 2 (free)**: import `MacroActions.json` from `TE macros/Conferences/New Stars of Data 2025/` via *Macros → Manage Macros → Import*.
3. **Semantic Link Labs notebooks**: upload the notebook to a Fabric workspace, attach a Lakehouse where needed, set the parameters and run all cells.
4. **Knowledge site**: see [Knowledge site](#knowledge-site) to preview or build it locally.
