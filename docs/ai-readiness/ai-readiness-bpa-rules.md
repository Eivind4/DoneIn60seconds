---
draft: true
publish: 2026-12-11
icon: lucide/shield-check
video_seconds: 45
te_versions: [TE3]
---

# Tailored BPA rules: AI readiness

![Tailored BPA rules: AI readiness](../assets/videos/ai-readiness-bpa-rules.mp4)

## The challenge

The standard BPA rules focus on performance and modeling, not on whether a model is ready for AI. Missing descriptions, unclear names and duplicate field names make the answers from Copilot less accurate.

## The solution

A set of BPA rules tailored for AI readiness flags the objects to fix, so every model is checked the same way.

!!! todo
    Add the script or notebook to `TE macros/TE3 shared/` (named with the day number), list it under `macros`, and add the `??? example` block.

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Tutorial: Prepare semantic model for AI](https://learn.microsoft.com/power-bi/create-reports/tutorial-copilot-power-bi-prepare-model), Microsoft Learn: what to simplify, such as duplicate field names and unused objects.
- [Best Practice Rules](https://github.com/microsoft/Analysis-Services/tree/master/BestPracticeRules), Microsoft on GitHub: the standard BPA rule set for Tabular Editor.
