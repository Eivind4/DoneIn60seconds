---
draft: true
publish: 2026-12-07
icon: lucide/chart-candlestick
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/07 - Error bars for actual vs target DaxUDF.cs
---

# Error bars for actual vs. target with DaxUDF

![Error bars for actual vs. target with DaxUDF](../assets/videos/error-bars.mp4)

## The challenge

Showing where the actual is clearly above or below the target makes a chart easier to read. Error bars can do this, but they need measures that only return a value when the deviation crosses a threshold.

!!! warning "Be aware"
    This is preparation for visuals. It may not be required if the content is consumed in a Fabric app.

## The solution

The macro creates two DaxUDF that return the actual only when it is more than a threshold above or below the target, and blank otherwise. Add the measures as error bars on the target series.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/07 - Error bars for actual vs target DaxUDF.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Use the Analytics pane in Power BI: error bars](https://learn.microsoft.com/power-bi/transform-model/desktop-analytics-pane#add-error-bars), Microsoft Learn: how to add error bars by field.
- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
