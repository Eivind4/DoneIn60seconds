---
draft: true
publish: 2026-12-15
icon: lucide/database-zap
video_seconds: 45
---

# Mini models with materialized lake views

![Mini models with materialized lake views](../assets/videos/mini-models-materialized-lake-views.mp4)

## The challenge

Smaller models often need their own, smaller tables, and keeping these tables in sync with the source takes work.

## The solution

Materialized lake views prepare the data for each smaller model in the lakehouse, and Fabric refreshes them in the right order.

!!! todo
    Add the script or notebook to `TE macros/TE3 shared/` (named with the day number), list it under `macros`, and add the `??? example` block.

## More on the solution

!!! todo "Your words"
    Design choices, limitations, alternatives and similar scripts.

## References

- [What are materialized lake views in Microsoft Fabric?](https://learn.microsoft.com/fabric/data-engineering/materialized-lake-views/overview-materialized-lake-view), Microsoft Learn: when to use them, and how refresh and dependencies work.
