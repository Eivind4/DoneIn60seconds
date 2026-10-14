---
draft: true
publish: 2026-12-13
icon: lucide/layers-2
complexity: Advanced
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/13 - Master model perspectives.cs
---

# Create model from perspective

![Create model from perspective](../assets/videos/create-model-from-perspective.mp4)

## The challenge

Maintaining many models with the same tables and measures means that the same change has to be made in several places.

## The solution

Develop one master model, and use perspectives to define the smaller models. The macro creates or updates one model per perspective, so a change is made once, in the master model.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/13 - Master model perspectives.cs"
    ```

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

!!! todo "Your words"
    - [Title](url): one line on why it is worth reading.
