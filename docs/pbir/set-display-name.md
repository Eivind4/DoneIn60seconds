---
draft: true
publish: 2026-12-19
icon: lucide/tags
complexity: Beginner
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/19 - Set display names.cs
  - TE macros/TE3 shared/19 - Batch rename display names.cs
---

# Set and rename display names

![Set and rename display names](../assets/videos/set-display-name.mp4)

## The challenge

Display names let a visual show a friendlier name than the field name, but setting them visual by visual is slow, and they drift apart across reports.

!!! warning "Be aware"
    The macro changes the PBIR files. The reports are often thin reports, live connected to the semantic model in the service, so make sure the model in the service has the same changes as the model you work on. Keep the reports in source control so the changes are easy to review and undo.

## The solution

The first macro sets display names for selected measures and columns across one or more reports. The second standardizes existing display names in batch, based on the underlying measure or column.

??? example "Show the script: Set display names"
    ```csharp
    --8<-- "TE macros/TE3 shared/19 - Set display names.cs"
    ```

??? example "Show the script: Batch rename display names"
    ```csharp
    --8<-- "TE macros/TE3 shared/19 - Batch rename display names.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
