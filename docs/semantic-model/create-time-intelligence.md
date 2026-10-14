---
draft: true
publish: 2026-12-02
icon: lucide/calendar-range
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/02 - Create time intelligence DaxUDF.cs
  - TE macros/TE3 shared/02 - Time intelligence calculation group.cs
---

# Create time intelligence DaxUDF

![Create time intelligence DaxUDF](../assets/videos/create-time-intelligence.mp4)

## The challenge

Time intelligence, such as YTD, previous year and rolling periods, is needed in almost every model. Writing it measure by measure takes time, and the logic ends up slightly different from model to model.

!!! warning "Be aware"
    DAX user-defined functions (DaxUDF) require compatibility level 1702 or higher.

## The solution

The macro generates the time intelligence calculations you select as DaxUDF, built on the calendar options from day 1. The functions are created once and reused by every measure (day 3).

!!! todo
    Example: how many days to Christmas.

??? example "Show the script: Create time intelligence DaxUDF"
    ```csharp
    --8<-- "TE macros/TE3 shared/02 - Create time intelligence DaxUDF.cs"
    ```

??? example "Show the script: Time intelligence calculation group"
    ```csharp
    --8<-- "TE macros/TE3 shared/02 - Time intelligence calculation group.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Introducing calendar-based time intelligence in DAX](https://www.sqlbi.com/articles/introducing-calendar-based-time-intelligence-in-dax/), SQLBI: the calendar-based time intelligence functions that the DaxUDFs build on.
- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
- [Hiding future dates for calculations in DAX](https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/), SQLBI: why and how to stop time intelligence after the last date with data.
