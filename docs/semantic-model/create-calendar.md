---
draft: true
publish: 2026-12-01
complexity: Beginner
video_seconds: 45
te_versions: [TE3]
macros:
  - TE macros/TE3 shared/01 - Create date table.cs
---

# Create a date table

![Create a date table](../assets/videos/create-calendar.mp4)

A good way to start a Christmas calendar… is a calendar.

The purpose is to set you up for the coming work, and to get you up and running fast for prototyping.

## The challenge

It is time-consuming to organize your date table:

- Organize in folders
- Sort month and text columns by numeric columns
- Add descriptions
- Mark as date table (unless using calendar options)
- Create hierarchies
- Hide unnecessary columns (e.g. those only used for sorting)
- Add calendar options (see [Tabular Editor](https://docs.tabulareditor.com/en/tutorials/calendars.html)
  and [SQLBI](https://www.sqlbi.com/articles/introducing-calendar-based-time-intelligence-in-dax/))
- Set up relationships

Copying the query will not include all of this.

Not everyone celebrates Christmas, and not everyone uses the same date table (different languages, etc.).

!!! warning "Be aware"
    The important thing is that the date table has calendar options. They are in preview at the time
    of writing, but they are recommended for new solutions, as they can be faster to run. See the
    [on-demand webinar from Tabular Editor](https://tabulareditor.com/resources/upcoming-and-on-demand-events/on-demand-webinar-calendar-based-time-intelligence).

## The solution

We can set up the table with a macro that dynamically sets the min and max dates based on your model,
marks it as a date table and adds relationships, in addition to solving the tedious setup of the date table.

The calendar is based on the [macro](https://github.com/Eivind4/DoneIn60seconds/blob/main/TE%20macros/TE3%20shared/01%20-%20Create%20date%20table%20%28table%20only%29.cs)
from Tabular Editor that creates a calculated table. This one creates it in Power Query.

The purpose of this macro is to get you started in an easy way and set you up for the following posts.
It adds the calendar options, which we will use later.

??? example "Show the script"
    ```csharp
    --8<-- "TE macros/TE3 shared/01 - Create date table.cs"
    ```

## More on the solution

Ideally, you have a central date table. The first time you set it up, it is not deterministic, and you
can use a [skill](https://github.com/Eivind4/DoneIn60seconds/blob/main/skills/date-table-formatting/SKILL.md).
Then reuse it with a TMDL script to add the date table.

Calendar options also have other use cases. They make more complex calculations easier, they can have
month granularity, and they make it easier to filter on day (which is why day is not mapped), e.g. for
finding YTD for a given day.

**Alternatives:** use [Bravo](https://bravo.bi/) to set up a calendar, in particular if you are not
using the Gregorian calendar. Note that it only sets up the date table, not the remaining settings.

## References

- [Calendars (Enhanced Time Intelligence)](https://docs.tabulareditor.com/en/tutorials/calendars.html),
  Tabular Editor: how to set up calendar options in Tabular Editor 3.
- [Introducing calendar-based time intelligence in DAX](https://www.sqlbi.com/articles/introducing-calendar-based-time-intelligence-in-dax/),
  SQLBI: how the calendar-based time intelligence functions work, including weekly and other non-Gregorian calendars.
- [An introduction to calendar-based time intelligence in Power BI – with Tabular Editor 3](https://tabulareditor.com/resources/upcoming-and-on-demand-events/on-demand-webinar-calendar-based-time-intelligence),
  on-demand webinar: why calendar options are worth using in new solutions.
