---
draft: true
publish: 2026-12-09
icon: lucide/calendar-x-2
video_seconds: 45
te_versions: [TE3]
---

# Detect and fix date inconsistencies with BPA

![Detect and fix date inconsistencies with BPA](../assets/videos/fix-inconsistent-dates.mp4)

## The challenge

Date columns easily end up set up differently across tables and models, for example with different format strings or data types. This is hard to spot by hand, and it shows up as inconsistent reports.

## The solution

A Best Practice Analyzer (BPA) rule detects the inconsistent date columns, and its fix expression corrects them in one go.

!!! todo
    Add the script or notebook to `TE macros/TE3 shared/` (named with the day number), list it under `macros`, and add the `??? example` block.

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Best Practice Rules](https://github.com/microsoft/Analysis-Services/tree/master/BestPracticeRules), Microsoft on GitHub: the standard BPA rule set for Tabular Editor.
- [Semantic Link Labs](https://github.com/microsoft/semantic-link-labs), Microsoft on GitHub: run BPA and other model checks from a Fabric notebook.
