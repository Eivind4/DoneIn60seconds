---
draft: true
publish: 2026-12-17
icon: lucide/replace-all
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/17 - Batch rename measures.cs
---

# Batch rename measures

![Batch rename measures](../assets/videos/batch-rename-measures.mp4)

## The challenge

Renaming many measures, for example to follow a naming convention, is slow one by one, and it breaks the reports.

!!! warning "Be aware"
    The macro changes the PBIR files. Save the report in Power BI Desktop first, and keep it in source control so the changes are easy to review and undo.

## The solution

The macro renames the selected measures in the model and updates the references in the PBIR files in the same run.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/17 - Batch rename measures.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [DAX naming conventions](https://docs.sqlbi.com/dax-style/dax-naming-conventions), SQLBI: a naming convention to rename towards.
- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
