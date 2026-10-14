---
name: date-table-formatting
version: 0.1.0
description: Format an existing Power BI/Analysis Services date/calendar table to production standard - descriptions, formatting, sort-by columns, display folders, Gregorian + Weekly calendar options, and fact-table relationships. Automatically invoke when the user asks to "format the date table", "format the calendar table", "clean up my date table", "add calendar options", "mark as date table and configure calendar", or wants to apply standard date-table conventions to a table in a semantic model.
---

# Date Table Formatting

Takes a raw/unformatted date or calendar table (already loaded with data via Power Query) and brings it up to a standard, production-ready state: correct metadata, sort-by columns, display folders, descriptions, Gregorian + Weekly calendar column groups, and relationships to fact tables.

**This skill assumes the table already contains the right columns and data.** It does not invent new date columns or write the M query from scratch - it formats what already exists. If the table is missing obvious building blocks (e.g. no ISO week columns at all and the user wants a Weekly calendar), tell the user what's missing rather than fabricating columns.

## Important calibration note

This skill was distilled from a session where an *already-formatted sibling date table existed in the same model* to copy folder names, sort-by targets, and calendar mappings from. That made the result look deceptively polished. **In a new model there usually is no reference table** - so this skill encodes explicit default rules below instead of "copy what's already there." Treat every rule as a sensible default, not gospel: state assumptions out loud and ask the user when a choice genuinely could go either way (see "Ask, don't guess" section).

## Tooling

Prefer the `te` CLI (Tabular Editor CLI) skill (`te-cli`) for connecting to and editing the model - it tolerates malformed/corrupted model metadata better than raw TOM, and its C# scripting (`c-sharp-scripting` skill) lets every change below be applied in a single transactional script.

- Local Power BI Desktop: `te connect --local` first; if that finds nothing, discover the port via the process's AS server (`netstat`/`Get-NetTCPConnection` correlated to `msmdsrv.exe`/PBIDesktop PID) and connect directly: `te connect localhost:<port>`.
- Every subsequent `te` command needs `-s localhost:<port> -d <databaseId>` explicitly - connection state does not persist across separate shell invocations.
- Do all edits in one `te script ... --save` C# script rather than many small `te set` calls, so the change is one atomic transaction and easy to review/undo.
- Validate afterwards with `te validate` (structural) and a quick DAX sanity query (row count, min/max date) via `te query`.
- If raw TOM/PowerShell reports 0 tables or empty collections on a model that clearly has data (visible via ADOMD DMVs), suspect a dangling annotation/orphaned object reference corrupting TOM's full deserialization; `te` is typically more tolerant and should be used instead of debugging the corruption unless the user asks for it to be fixed.

## Workflow

### 1. Identify the target table and gather facts

- Confirm with the user (or infer unambiguously) which table is the date table to format, and which fact table(s) it should relate to.
- List the table's columns (`te ls "<Table>" --output-format json`) and classify each: numeric/technical (Year, Month no, Quarter no, Date_Key, ISO week, is_Weekend, etc.) vs. text/label (Month name, Day name, Year quarter, etc.) vs. the primary `Date` column.
- Check whether ISO week-related columns exist (ISO year, ISO week, ISO year week, or equivalents) - this determines whether a Weekly calendar option is possible.
- Look at the fact table(s) for date-typed foreign key columns (e.g. Order Date, Delivery Date, Ship Date) to plan relationships.

### 2. Table and column metadata

Apply to the date table itself:
- `DataCategory = "Time"` (marks it as a date table candidate for time-intelligence)
- A short table `Description`

Apply to every column:
- **Description** in the user's preferred language - concise, one sentence, states what the column represents (and format/derivation for computed slicer text like "Today"/"Current Year").
- **Do not summarize numeric columns**: set `SummarizeBy = None` on every numeric column (Year, Month no, Quarter no, Date_Key, ISO year, ISO week, is_Weekend, is_History, etc.). Numeric-looking flag/key columns should never default to Sum in visuals.
- **Format the primary Date column**: set `FormatString` (default `"MM/dd/yyyy"` unless the user/org has a documented date format standard - ask if unsure).
- **Hide technical/key columns** not meant for end users (e.g. `Date_Key`, sort-by helper columns if they aren't otherwise useful).
- **Display folders**: group into a small numbered set so folder order matches logical hierarchy, e.g.:
  `1. Slicer`, `2. Year`, `3. Quarter`, `4. Month`, `5. Week`, `6. Date`, `7. Day`, `8. Boolean`
  Don't invent more folders than needed - a handful of clearly-named ones beats one per column. If the model already has a folder-naming convention elsewhere, match it instead of the default above.

### 3. Sort-by columns

Rule: **every non-numeric/text column that has a natural numeric or chronological order must get a `SortByColumn`** pointing at the numeric column that encodes that order. Numeric columns themselves don't need a sort-by (they already sort correctly).

Typical mappings (adapt names to what actually exists in the table):

| Text/label column | Sort by |
|---|---|
| Date slicer | Date |
| Year slicer | Year |
| Quarter (e.g. "Q1") | Quarter no |
| Month name / Month short name | Month no |
| Year month / Month name Year / Year month slicer | Year month no |
| Day name | Day no of week |
| Week slicer | ISO year week (if already zero-padded/sortable) or a dedicated numeric week-sort column |

If no numeric column exists that encodes the right order (e.g. "Year quarter" as `"2024 Q1"` sorts correctly lexically because it's zero-padded/fixed-width), it's fine to leave it on default sort - note this decision rather than silently skipping it.

### 4. Calendar options (Gregorian + Weekly)

Create both as separate named calendar column groups on the table (TOMWrapper: `Table.AddCalendar(name)` then `Calendar.AddTimeUnit(TimeUnit, primaryColumn, ...associatedColumns)`). Do not include `DayOfWeek` or `DayOfMonth` as mapped time units in either calendar - those are excluded by convention.

**Gregorian** (only include time units whose columns exist):
- Year -> Year (associated: Year slicer)
- Quarter -> Year quarter
- QuarterOfYear -> Quarter no (associated: Quarter)
- Month -> Year month no (associated: Year month, Year month slicer, Month name Year)
- MonthOfYear -> Month no (associated: Month name, Month short name)
- Date -> Date (associated: Date slicer, Date_Key)

**Weekly** (only if ISO week columns exist):
- Year -> ISO year
- Week -> ISO year week (associated: Week slicer)
- WeekOfYear -> ISO week
- Date -> Date (associated: Date slicer, Date_Key)

**Calendar names must be unique model-wide, not just per-table.** If a name collides with one already used elsewhere in the model, TE silently auto-suffixes it (e.g. "Gregorian 1") rather than erroring - avoid this by giving explicit unique names, e.g. `"Gregorian (<Table Name>)"` / `"Weekly (<Table Name>)"`, unless the user prefers plain "Gregorian"/"Weekly" and there's no naming conflict.

### 5. Relationships to fact table(s)

- For each date-typed foreign key column on a fact table, create a relationship to the date table's `Date` column: many-to-one, single-direction (`CrossFilteringBehavior.OneDirection`).
- **Ask, don't guess, which relationship should be active** when a fact table has multiple date roles (e.g. Order Date vs. Delivery Date/Ship Date) - only one relationship per fact table can be active at a time to the same dimension unless using inactive relationships + `USERELATIONSHIP`. Default assumption (state it, but confirm): the primary transactional date (e.g. Order Date) is active, secondary date roles (Delivery/Ship Date) are inactive.
- Two different, mutually-unconnected date tables can each have an active relationship from the same fact column without conflict (no ambiguity/loop) - this is valid and doesn't need to be avoided if the model has multiple date tables.

### 6. Power Query / M review

- Run the model's/skill's M formatter (e.g. `te format`) for consistent indentation and structure.
- Check whether the M expression queries an external source (SQL, Lakehouse, etc.) or generates dates purely in-memory (e.g. `List.Dates(...)`). **Query folding only applies when there's an external source to fold into** - if the date table is self-generated, say so explicitly instead of forcing artificial folding changes.
- If there is an external source, apply the general Power Query folding guidance from the `power-query` skill (filter/select/type-transform early, custom columns last).

### 7. Ask, don't guess

Points that genuinely vary by organization/model and should be a quick clarifying question rather than an assumption:
- Description language
- Date format string (if there's no established org standard)
- Which fact-table relationship should be active vs. inactive when multiple date roles exist
- Calendar column-group naming, if it would collide with existing names
- Display folder scheme, if the model already has an established convention elsewhere

### 8. Validate

- `te validate` for structural errors.
- A quick DAX sanity check: row count, min/max date, and (if relationships were added) a `CALCULATE`/`USERELATIONSHIP` test that both active and inactive relationships resolve as expected.
- Summarize every change made (table, columns, calendars, relationships) back to the user before finishing.

## Related skills

- `te-cli` - driving the `te` CLI itself
- `c-sharp-scripting` - writing the TOMWrapper C# script that applies all changes in one transaction
- `power-query` - M authoring/folding guidance if the date table's source is external
