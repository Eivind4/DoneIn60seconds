---
draft: true
video_seconds: 30
---

# C# scripting in Tabular Editor

Tabular Editor comes in a free version, Tabular Editor 2, and a licensed version, Tabular Editor 3.
Both have the C# scripting functionality, and the challenges in this series can be solved with both tools.
There are some differences, though:

!!! todo "Your words, with suggestions from Claude to accept, rewrite or delete"
    Your notes: responsive input boxes; can add the Roslyn compiler.

    | | Tabular Editor 2 | Tabular Editor 3 |
    |---|---|---|
    | License | Free, open source | Paid |
    | C# version | Older compiler by default; newer syntax (e.g. string interpolation) needs the [Roslyn compiler](https://docs.tabulareditor.com/te2/Advanced-Scripting.html), set under File > Preferences. For advanced users. | Modern C# out of the box |
    | Script editor | Basic | Code completion, syntax highlighting, script recorder |
    | Saved scripts | Custom actions (`CustomActions.json`) | Macros (`MacroActions.json`), can be added to toolbars and menus |
    | Output and errors | Pop-up dialogs | Messages window |

The scripts in this series are created for Tabular Editor 3, as it saves my (and your) valuable time.
If you are limited to Tabular Editor 2, your AI agent can help you make a script compile and run in
Tabular Editor 2 as well. Since I have created some scripts for Tabular Editor 2 myself, I have made a skill
that may help you get started.

!!! todo
    Attach the Tabular Editor 2 skill and link it here.

Many of the scripts were created while preparing for New Stars of Data in May 2025, and have since been
improved with DAX UDF functionality. For this, I used the plugin skills from Data Goblins, which I recommend:

- [data-goblin/power-bi-agentic-development](https://github.com/data-goblin/power-bi-agentic-development):
  a plugin marketplace of Power BI skills, subagents and hooks for Claude Code and GitHub Copilot.
- [C# scripting skill](https://github.com/data-goblin/power-bi-agentic-development/tree/main/plugins/tabular-editor/skills/c-sharp-scripting)
  in the Tabular Editor plugin.

## A few things to help you

- Tabular Editor 3 has a Messages window, where you can copy errors and give them as input to AI.
  In Tabular Editor 2, you can use screenshots.
- With the September 2026 release, you can access Tabular Editor through an MCP server from your favourite
  AI agent. I have not tested it for C# scripting at the time of writing.

## Add and save a script as a macro

![Add and save a script as a macro](../assets/videos/tabular-editor.mp4)

!!! todo
    Record a short video showing how to add and save a script.

More details: [Creating macros | Tabular Editor Documentation](https://docs.tabulareditor.com/en/features/creating-macros.html)

## Other C# references

- Tabular Editor: [Beginner C# Scripts](https://docs.tabulareditor.com/common/CSharpScripts/csharp-script-library-beginner.html)
- Bernat Agulló: [TabularEditor-Scripts / Basic](https://github.com/bernatagulloesbrina/TabularEditor-Scripts/tree/main/Basic)
- Alexander Korn: [PBI-Tools / PBI Toolbox - Tabular Editor Macros](https://github.com/KornAlexander/PBI-Tools/tree/main/PBI%20Toolbox%20-%20Tabular%20Editor%20Macros)
- PowerBI.tips: [TabularEditor-Scripts](https://github.com/PowerBI-tips/TabularEditor-Scripts):
  community-driven scripts for Tabular Editor 2 and 3.
- Michael Kovalsky: [Tabular](https://github.com/m-kovalsky/Tabular):
  useful code for tabular modeling and automation.
