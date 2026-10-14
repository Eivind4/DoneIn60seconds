---
draft: true
publish: 2026-12-08
icon: lucide/palette
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/08 - Conditional formatting by thresholds DaxUDF.cs
  - TE macros/TE3 shared/08 - Apply conditional formatting to visuals.cs
---

# Conditional formatting by thresholds with DaxUDF

![Conditional formatting by thresholds with DaxUDF](../assets/videos/conditional-formatting.mp4)

## The challenge

Coloring values as good, bad or neutral compared with a target is a common requirement. Setting up rules in each visual is slow, and the rules easily end up different across visuals and reports.

!!! warning "Be aware"
    This is preparation for visuals. It may not be required if the content is consumed in a Fabric app.

## The solution

The first macro creates DaxUDF that return a color, or a good/bad/neutral flag, based on a measure (typically the % difference to a target) and an upper and lower bound, both where an increase is good and where a decrease is good. The second macro applies the color measure as conditional formatting to visuals in the PBIR files.

??? example "Show the script: Conditional formatting by thresholds DaxUDF"
    ```csharp
    --8<-- "TE macros/TE3 shared/08 - Conditional formatting by thresholds DaxUDF.cs"
    ```

??? example "Show the script: Apply conditional formatting to visuals"
    ```csharp
    --8<-- "TE macros/TE3 shared/08 - Apply conditional formatting to visuals.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [Conditional formatting in Power BI visuals](https://learn.microsoft.com/power-bi/visuals/power-bi-visualization-conditional-formatting), Microsoft Learn: which visuals support it, and how to use a color measure as field value.
- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
- [Power BI Desktop project report folder](https://learn.microsoft.com/power-bi/developer/projects/projects-report#pbir-format), Microsoft Learn: the PBIR files that the macros read and change.
