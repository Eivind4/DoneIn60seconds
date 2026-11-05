# Issue drafts: model cleaning and setup flow

Drafts to create as GitHub issues. Issue 0 is the tracking issue; issues 1-12 are its sub-issues.
Delete this file once the issues exist.

---

## 0. Model cleaning and setup flow (tracking)

Macros that clean up a model at the end of development (typically fixing Best Practice Rules),
set up new models (greenfield), check existing ones (brownfield), and organize the macros themselves.

Cleaning order:

1. Remove unused items.
2. Run the macros that fix BPA rules.
3. Format measures.
4. Do not summarize columns, display folders and model documentation.

---

## 1. Cleaning flow: remove items before fixing BPA rules

Run the cleanup in a fixed order: remove unused items first, then run the macros that fix BPA rules,
so the fixes do not touch objects that are about to be deleted.

- Starting point: `TE macros/TE3/16 - Classify model dependencies from PBIR.cs` (detect unused objects).
- BPA rules: `TE macros/Conferences/New Stars of Data 2025/9. Load best practice rules.csx`.
- Site page: `docs/best-practice-rules/automate-fixes.md`.

## 2. Format measures

Include formatting of measures in the cleaning flow (DAX formatting and format strings).

- Starting point: `TE macros/Conferences/New Stars of Data 2025/9. Format measures - Max 20.csx`.

## 3. Do not summarize columns

Macro that sets columns to "Do not summarize".

- Starting point: `TE macros/Conferences/New Stars of Data 2025/9. Do not summarize numeric columns.csx`.

## 4. Display folders

Macro that organizes measures and columns into display folders.

## 5. Model documentation

Macro that adds model documentation.

- Starting point: `TE macros/Conferences/New Stars of Data 2025/8. Create model documentation.csx`.
- Later: decide and describe what needs to be documented.

## 6. Greenfield: setup walkthrough

For a new model, let the user walk through what to set, either for the environment (first time)
or for the model. For example:

- Add currency conversion?
- Add model documentation?
- Add report perspectives?

## 7. Brownfield: checks with reports

For an existing model, add checks together with the reports that use it.

## 8. Report perspectives: update after report changes

If report perspectives are created, update the perspective when the report is updated.

- Related: `TE macros/TE3/18 - Master model - create or update perspective models.cs`.

## 9. Model performance scan (Kovalsky macros)

Add a model performance scan based on the macros from Michael Kovalsky. Check for:

- High cardinality columns.
- High cardinality in relationships.
- Referential integrity (may be difficult to fix from here alone).

## 10. Same measure name, different DAX across models

Check whether any measure has the same name as a measure in another model in the repository,
but a different DAX expression.

- If it is on purpose, accept it and add an annotation that excludes it from the next check.
- Related page: `docs/best-practice-rules/duplicate-measures.md`.

## 11. Same DAX expression, different measure name across models

Check whether a measure has the same expression as a measure in another model in the repository,
but a different name.

## 12. Macro actions for source control: numbering, naming and folders

Create a macro actions file (`MacroActions.json`) for source control that includes the previous
macros, with organized numbers and naming.

- Consider a folder structure that follows the organization of the macros.
- Starting point: `TE macros/Conferences/New Stars of Data 2025/MacroActions.json`.
