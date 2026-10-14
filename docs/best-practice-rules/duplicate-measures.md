---
draft: true
publish: 2026-12-23
icon: lucide/copy-x
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/23 - Replace duplicate measures in visuals.cs
---

# Detect and fix duplicate measures

![Detect and fix duplicate measures](../assets/videos/duplicate-measures.mp4)

## The challenge

Over time, models get several measures with the same logic under different names. Duplicates confuse users and AI, and a change has to be made in several places.

!!! warning "Be aware"
    The macro changes the PBIR files. Save the report in Power BI Desktop first, and keep it in source control so the changes are easy to review and undo.

## The solution

A BPA rule detects measures with the same expression. The macro replaces the duplicate with the measure you keep in the report visuals, so the duplicate can be deleted safely.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/23 - Replace duplicate measures in visuals.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Best Practice Rules](https://github.com/microsoft/Analysis-Services/tree/master/BestPracticeRules), Microsoft on GitHub: the standard BPA rule set for Tabular Editor.
- [Tutorial: Prepare semantic model for AI](https://learn.microsoft.com/power-bi/create-reports/tutorial-copilot-power-bi-prepare-model), Microsoft Learn: what to simplify, such as duplicate field names and unused objects.
