---
draft: true
publish: 2026-12-06
icon: lucide/trophy
complexity: Beginner
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/06 - Rank by selected columns DaxUDF.cs
---

# Rank by selected columns with DaxUDF

![Rank by selected columns with DaxUDF](../assets/videos/ranking.mp4)

## The challenge

Ranking a measure by customer, product or another column is a common need, for example for top N lists. Ranking in DAX is easy to get slightly wrong, especially for totals and ties.

## The solution

Select a base measure and the columns to rank by, and run the macro. It creates the ranking functions for each column, based on the patterns from DAX Patterns. Only columns from tables that work as ranking dimensions can be selected.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/06 - Rank by selected columns DaxUDF.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.
    Similar scripts in the private repo: TopN measures, Pareto analysis DaxUDF.

## References

- [Ranking](https://www.daxpatterns.com/ranking/), DAX Patterns: the ranking patterns the functions are based on.
- [RANK](https://learn.microsoft.com/dax/rank-function-dax), Microsoft Learn: the window function for ranking, including how ties are handled.
- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
