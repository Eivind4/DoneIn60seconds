---
draft: true
publish: 2026-12-05
icon: lucide/git-compare
complexity: Beginner
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/05 - Compare actual and target DaxUDF.cs
---

# Compare actual and target with DaxUDF

![Compare actual and target with DaxUDF](../assets/videos/comparison.mp4)

## The challenge

Comparing an actual with a target, budget or previous period is one of the most common needs in a report: the deviation, the % change or the % of total. Each comparison is a new measure, written again for each pair.

## The solution

Select the actual measure first and the target second, and run the macro. It adds the comparison functions (deviation, % change and % of total) if they are missing, and creates a measure for each function you pick.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/05 - Compare actual and target DaxUDF.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
