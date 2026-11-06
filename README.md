# DoneIn60seconds

Scripts and resources from various *Done in 60 Seconds* community sessions on Power BI semantic model automation with Tabular Editor and Semantic Link Labs. The content covers time intelligence, model formatting, best-practice analysis, and DAX User-Defined Functions (UDFs).

---

## Repository structure

```
TE macros/
  Conferences/
    New Stars of Data 2025/     ← session scripts (TE2 compatible)
    Data Toboggan - Summer 2025/ ← session scripts + presentation
  TE2/                           ← standalone TE2 scripts
  TE3/                           ← TE3-only scripts (DAX UDFs, Calendars)
Dax UDF/                         ← TMDL snippets for DAX UDF patterns
Semantic Link Labs/
  BPA/                           ← Fabric notebook for BPA with SLL
  BPA-SemanticLinkLabsReport/    ← PBIR report for BPA results
```

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

## TE macros — TE3

Scripts that require **Tabular Editor 3** and make use of TE3-exclusive features such as DAX UDFs and named Calendars.

| File | Purpose |
|------|---------|
| `Generate Time Intelligence DaxUDF.cs` | Interactive script that creates a full set of time intelligence DAX UDFs in the model. Supports both standard date-column mode and TE3 Calendar mode (fiscal/weekly). Each UDF receives `FunctionCategory`, `MeasureSuffix`, and `FolderSuffix` annotations so a companion macro can auto-generate measures for any selected base measure. |
| `DaxUDF - Create measure from OneParameter functions.cs` | Creates measures by calling one-parameter DAX UDFs on selected base measures, reading the `MeasureSuffix` and `FolderSuffix` annotations set by the generator above. |
| `DaxUDF - Compare first selected measure with second selected measure.cs` | Creates a comparison UDF measure between two selected measures. |
| `Comparison Function.tmdl` | TMDL snippet for a reusable comparison DAX UDF. |

## TE macros — TE2

| File | Purpose |
|------|---------|
| `Create Time Intelligence DaxUDF.csx` | TE2-compatible version of the DAX UDF generator (subset of calculations). |

---

## Dax UDF

TMDL reference snippets for DAX UDF patterns used across scripts.

| File | Purpose |
|------|---------|
| `Comparison Function.tmdl` | TMDL definition of a comparison DAX UDF (difference and % difference between two measures). |

---

## Semantic Link Labs

Resources for running Best Practice Analyzer (BPA) analysis at scale using the [Semantic Link Labs](https://github.com/microsoft/semantic-link-labs) Python library inside Microsoft Fabric.

| File | Purpose |
|------|---------|
| `BPA/Fabric Februay - SLL BPA demo Shared.ipynb` | Fabric notebook demonstrated at the Fabric February community event. Runs BPA rules across one or more semantic models and writes results to a Lakehouse table. |
| `BPA-SemanticLinkLabsReport/` | Power BI Enhanced Report (PBIR) for visualising the BPA results produced by the notebook above. |

---

## Getting started

1. **Tabular Editor 2 (free)** — import `MacroActions.json` from `TE macros/Conferences/New Stars of Data 2025/` via *Macros → Manage Macros → Import*.
2. **Tabular Editor 3** — open any `.cs` / `.csx` file from `TE macros/TE3/` in the TE3 script editor and run directly, or save as a macro.
3. **Semantic Link Labs notebook** — upload `Fabric Februay - SLL BPA demo Shared.ipynb` to a Fabric workspace, attach a Lakehouse, and run all cells.
