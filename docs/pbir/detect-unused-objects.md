---
draft: true
publish: 2026-12-20
icon: lucide/network
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/20 - Classify model dependencies.cs
---

# Classify model dependencies: report dependency

![Classify model dependencies: report dependency](../assets/videos/detect-unused-objects.mp4)

## The challenge

Before changing or cleaning up a model, you need to know which tables, columns and measures are used in reports.

## The solution

The macro reads the PBIR files of the reports and classifies each model object by whether, and where, it is used.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/20 - Classify model dependencies.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
