---
draft: true
publish: 2026-12-03
icon: lucide/sigma
complexity: Intermediate
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/03 - Create measures from model dependent DaxUDF.cs
---

# Create measures from model dependent DaxUDF

![Create measures from model dependent DaxUDF](../assets/videos/measures-with-time-intelligence.mp4)

## The challenge

When the functions exist, you still need one measure for each combination of base measure and function, each with a name, display folder, format string and description.

## The solution

Select one or more base measures, run the macro and pick the functions to apply. It creates one measure per combination, with name, display folder, format string and description taken from the base measure and the function.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/03 - Create measures from model dependent DaxUDF.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [DAX user-defined functions](https://learn.microsoft.com/dax/best-practices/dax-user-defined-functions), Microsoft Learn: syntax, parameter types and limitations of DaxUDF.
