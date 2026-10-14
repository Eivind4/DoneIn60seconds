---
draft: true
publish: 2026-12-04
icon: lucide/sliders-horizontal
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/04 - Disconnected date table.cs
  - TE macros/TE3 shared/04 - Dynamic axis from slicer selection.cs
---

# Dynamic axis from slicer selection

![Dynamic axis from slicer selection](../assets/videos/dynamic-axis.mp4)

## The challenge

Users often want to select one month in a slicer and see the months leading up to it on the axis, for example the last 6 or 12 months. A slicer on the date table filters the axis down to the selected month only.

!!! warning "Be aware"
    This is preparation for visuals. It may not be required if the content is consumed in a Fabric app.

## The solution

The first macro creates a disconnected copy of the date table to use in the slicer. The second creates a "Period to show" calculation group, based on DaxUDF, that shows the selected number of periods up to the selected date on the axis.

??? example "Show the script: Disconnected date table"
    ```csharp
    --8<-- "TE macros/TE3 shared/04 - Disconnected date table.cs"
    ```

??? example "Show the script: Dynamic axis from slicer selection"
    ```csharp
    --8<-- "TE macros/TE3 shared/04 - Dynamic axis from slicer selection.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Show previous 6 months of data from single slicer selection](https://www.sqlbi.com/articles/show-previous-6-months-of-data-from-single-slicer-selection/), SQLBI: the pattern behind the calculation group.
- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
