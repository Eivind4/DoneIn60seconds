---
draft: true
publish: 2026-12-18
icon: lucide/wrench
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/18 - Detect and repair broken visuals.cs
---

# Detect and repair broken visuals

![Detect and repair broken visuals](../assets/videos/fix-broken-measures.mp4)

## The challenge

When fields are renamed or removed, visuals break, and it is not obvious which reports and pages are affected.

!!! warning "Be aware"
    The macro changes the PBIR files. The reports are often thin reports, live connected to the semantic model in the service, so make sure the model in the service has the same changes as the model you work on. Keep the reports in source control so the changes are easy to review and undo.

## The solution

The macro scans the PBIR files for table, column and measure references that do not exist in the model, and lets you repair them.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/18 - Detect and repair broken visuals.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
