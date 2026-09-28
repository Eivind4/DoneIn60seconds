---
draft: true
hide:
  - navigation
---

# Done in 60 seconds

**Solving common Power BI challenges in less than 60 seconds.**

*Deterministic · token free · model-independent*

Development in Power BI can be time-intensive. The latest developments in AI can help, but:

- AI is non-deterministic by default
  ([A practical introduction to LLMs for data professionals](https://tabulareditor.com/blog/practical-introduction-to-llms-for-data-professionals)).
- With AI, you typically spend more time on planning and testing
  ([How to get started with agentic development for business intelligence and analytics](https://tabulareditor.com/blog/how-to-get-started-with-agentic-development-for-business-intelligence)).
- There are more efficient options that do not require prompting at all. That is what this series is about.

Tabular Editor is a tool for creating semantic models more efficiently and with better quality.
It has many features that help you in your daily work with semantic models, but this series focuses on only one of them:

- Using C# scripts for common challenges, to ensure a deterministic outcome.

Or perhaps more precisely: using AI to create reusable C# scripts. AI is great at C# scripting; I am not.
What you need is to know what is possible to solve, and a good understanding of the metadata.
You gain that by using Tabular Editor or TMDL, where you can see the properties, including some that are
hidden in the Power BI user interface.

The series does not explain general concepts, but links to articles where applicable. The goal of the series
is not only to solve common challenges in Power BI, but to solve them within 60 seconds, in every model.

<div class="grid cards" markdown>

-   :material-book-open-variant: **General**

    ---

    About, Tabular Editor, PBIR files, theme files and Semantic Link Labs.

    [:octicons-arrow-right-24: Start here](general/index.md)

-   :material-database-cog: **Semantic model**

    ---

    Calendar, time intelligence, comparison, ranking and more.

    [:octicons-arrow-right-24: Browse](semantic-model/index.md)

-   :material-file-tree: **PBIR**

    ---

    Rename, repair and clean up report references.

    [:octicons-arrow-right-24: Browse](pbir/index.md)

-   :material-check-decagram: **Best Practice Rules**

    ---

    Automate fixes for Best Practice Analyzer findings.

    [:octicons-arrow-right-24: Browse](best-practice-rules/index.md)

-   :material-sitemap: **Master model**

    ---

    Create smaller models from one master model.

    [:octicons-arrow-right-24: Browse](master-model/index.md)

</div>

More about me: [About](general/about.md)
