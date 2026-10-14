---
draft: true
publish: 2026-12-16
icon: lucide/pencil-line
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/16 - Rename tables and measures.cs
---

# Rename tables and measures

![Rename tables and measures](../assets/videos/update-column-measure-name.mp4)

## The challenge

Renaming a table, column or measure in the model breaks the visuals that use it in the reports built on the model.

!!! warning "Be aware"
    The macro changes the PBIR files. Save the report in Power BI Desktop first, and keep it in source control so the changes are easy to review and undo.

## The solution

The macro replaces the old table, column and measure references with the new ones in the PBIR files of the selected reports, so the reports keep working after the rename.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/16 - Rename tables and measures.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
