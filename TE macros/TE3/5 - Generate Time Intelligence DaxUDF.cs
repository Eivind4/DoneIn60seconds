// ADD TIME INTELLIGENCE AS DAX UDFs
// ===================================
//
// Purpose: Generate a set of reusable DAX User Defined Functions (UDFs) for common time-intelligence
// calculations across Accumulated, Moving total and averages, Semi-Additive, and Date-filtering
// categories - including YTD/PYTD, rolling N patterns, transaction-date-aware daily averages,
// RNm LY variants, IB/LastDate/open-period patterns, LastDate/Previous per Category variants,
// simple historical/remaining filters, and optional "Hide future dates" (HF) and/or
// Fiscal Year variants (YTD_FY / PYTD_FY). Each selected calculation is generated as its own DAX UDF
// (not a calculation group), named "<prefix>.TI_<suffix>"
// using the naming convention from https://docs.sqlbi.com/dax-style/dax-naming-conventions, so it can
// be wrapped by a measure per the "5 - Create measure from model dependent functions.cs" macro.
//
// Safety checks performed before creating each UDF:
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function's expression is reused/kept as-is instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one, Overwrite the
//       existing (outdated) definition, or Skip creating it (never silently overwritten).
//
#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualBasic;

Application.UseWaitCursor = false;
WaitFormVisible = false;

// ===================== DIALOG 1: Select Date source =====================
Func<Table, bool> IsDisconnected = t => {
    if (t == null) return false;
    var rels = Model.Relationships.Where(r => r.FromTable == t || r.ToTable == t).ToList();
    return rels.Count > 0 && rels.All(r => !r.IsActive);
};

// Prefer keyed DateTime columns on tables formally "Marked as a Date Table" (DataCategory ==
// "Time") when at least one exists in the model - this avoids accidentally listing an unrelated
// fact table's DateTime key column. If no table is marked as a Date Table, fall back to any
// connected keyed DateTime column (so the picker still works for models that never bothered to
// mark a date table).
var markedDateColumns = Model.AllColumns
    .Where(c => c.DataType == DataType.DateTime && c.IsKey && !IsDisconnected(c.Table) &&
                c.Table.DataCategory == "Time")
    .ToList();
var dateColumns = markedDateColumns.Count > 0
    ? markedDateColumns
    : Model.AllColumns
        .Where(c => c.DataType == DataType.DateTime && c.IsKey && !IsDisconnected(c.Table))
        .ToList();
var calendars = Model.AllCalendars
    .Where(c => !IsDisconnected(Model.Tables.FirstOrDefault(t => t.Calendars.Contains(c))))
    .ToList();
if (dateColumns.Count + calendars.Count == 0) {
    Info("Model contains no Date columns or Calendars.");
    return;
}

var dateForm = new Form
{
    Text = "Select Date source",
    MaximizeBox = true,
    MinimizeBox = false,
    ShowIcon = false,
    Size = new Size(
        Math.Min(760, Screen.PrimaryScreen.WorkingArea.Width - 100),
        Math.Min(360, Screen.PrimaryScreen.WorkingArea.Height - 100)),
    MinimumSize = new Size(540, 300),
    FormBorderStyle = FormBorderStyle.Sizable,
    StartPosition = FormStartPosition.CenterScreen
};
var dateFlow = new FlowLayoutPanel {
    Dock = DockStyle.Fill, Padding = new Padding(12), FlowDirection = FlowDirection.TopDown,
    WrapContents = false, AutoScroll = true
};
dateForm.Controls.Add(dateFlow);
Action<string> dateLabel = s => dateFlow.Controls.Add(new Label { Margin = new Padding(0,10,0,5), Text = s, AutoSize = true });

dateLabel("Choose Date column or calendar:");
var dateColumnSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 680 };
dateColumns.ForEach(c => dateColumnSelector.Items.Add(new ComboBoxItem(c)));
calendars.ForEach(c   => dateColumnSelector.Items.Add(new ComboBoxItem(c)));
dateColumnSelector.SelectedIndex = 0;
dateFlow.Controls.Add(dateColumnSelector);

dateLabel("Choose actual Date column (used inside VALUES for rolling avg, must be a column):");
var actualDateColLabel    = dateFlow.Controls[dateFlow.Controls.Count - 1] as Label;
var actualDateColSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 680 };
dateColumns.ForEach(c => actualDateColSelector.Items.Add(new ComboBoxItem(c)));
if (actualDateColSelector.Items.Count > 0) actualDateColSelector.SelectedIndex = 0;
dateFlow.Controls.Add(actualDateColSelector);

Action updateDateVis = () => {
    bool isCal = (dateColumnSelector.SelectedItem as ComboBoxItem)?.Calendar != null;
    actualDateColLabel.Visible    = isCal;
    actualDateColSelector.Visible = isCal;
};
dateColumnSelector.SelectedIndexChanged += (s, e) => updateDateVis();
    dateFlow.ClientSizeChanged += (s, e) => {
        int inputWidth = Math.Max(460, dateFlow.ClientSize.Width - 35);
        dateColumnSelector.Width = inputWidth;
        actualDateColSelector.Width = inputWidth;
    };
    updateDateVis();

var dateBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var dateOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var dateCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
dateBtnPanel.Controls.Add(dateOk);
dateBtnPanel.Controls.Add(dateCancel);
dateFlow.Controls.Add(dateBtnPanel);
dateForm.AcceptButton = dateOk;
dateForm.CancelButton = dateCancel;

if (dateForm.ShowDialog() != DialogResult.OK) { Info("Cancelled - no UDFs were created."); return; }

var selItem     = dateColumnSelector.SelectedItem as ComboBoxItem;
var cal         = selItem.Calendar;
bool isCalendar = cal != null;

var dateColumn = isCalendar ? cal.DaxObjectFullName : selItem.Column.DaxObjectFullName;

// actualDateCol is always a real column — used inside VALUES() and AVERAGEX()
var actualDateCol = isCalendar
                      ? (actualDateColSelector.SelectedItem as ComboBoxItem).Column.DaxObjectFullName
                      : selItem.Column.DaxObjectFullName;

// The table that owns the date column / calendar - used by the follow-up "Hide future dates" dialog
var dateTable = isCalendar
                  ? Model.Tables.FirstOrDefault(t => t.Calendars.Contains(cal))
                  : selItem.Column.Table;
var dateTableRef = dateTable == null ? null : $"'{dateTable.Name.Replace("'", "''")}'";
var factDateColumns = Model.AllColumns
    .Where(c => c.DataType == DataType.DateTime && c.Table != dateTable && !IsDisconnected(c.Table))
    .ToList();

bool isWeekly   = cal != null && cal.GetTimeUnits().Any(t =>
                      t.TimeUnit is TimeUnit.Week or TimeUnit.WeekOfMonth
                                 or TimeUnit.WeekOfYear or TimeUnit.WeekOfQuarter);
bool hasMonth   = cal == null || cal.GetTimeUnits().Any(t =>
                      t.TimeUnit is TimeUnit.Month or TimeUnit.MonthOfQuarter or TimeUnit.MonthOfYear);
bool hasQuarter = cal == null || cal.GetTimeUnits().Any(t =>
                      t.TimeUnit is TimeUnit.Quarter or TimeUnit.QuarterOfYear);

// ===================== DIALOG 2: Select calculations =====================
var aggForm = new Form
{
    Text = "Add Time Intelligence UDFs",
    MaximizeBox = true,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = false,
    Width = Math.Min(1050, Screen.PrimaryScreen.WorkingArea.Width - 80),
    Height = Math.Min(920, Math.Max(700, Screen.PrimaryScreen.WorkingArea.Height - 80)),
    MinimumSize = new Size(800, 650),
    FormBorderStyle = FormBorderStyle.Sizable,
    StartPosition = FormStartPosition.CenterScreen
};

var scrollPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
var flowPanel = new FlowLayoutPanel
{
    Padding = new Padding(10),
    FlowDirection = FlowDirection.TopDown,
    WrapContents = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink
};
scrollPanel.Controls.Add(flowPanel);
aggForm.Controls.Add(scrollPanel);

Action<string> AddLabel = s => {
    var label = new Label { Margin = new Padding(0,10,0,5), Text = s, AutoSize = true };
    flowPanel.Controls.Add(label);
};
Action<string> AddSectionLabel = s => {
    var label = new Label { Margin = new Padding(0,14,0,5), Text = s, AutoSize = true, Font = new Font(aggForm.Font, FontStyle.Bold) };
    flowPanel.Controls.Add(label);
};

var checks           = new List<CheckBox>();
var suffixes         = new List<TextBox>();
var descriptions     = new List<string>();
var hideFutureChecks = new List<CheckBox>();
var comparisonChecks = new List<CheckBox>();
var categories       = new List<string>();
var calcPanels       = new List<Panel>();
var categoryOrder    = new List<string> { "Accumulated", "Moving total and averages", "Semi-Additive", "Date-filtering" };

Action<string, string, string, string> AddCalc = (category, label, defaultSuffix, description) => {
    var panel    = new Panel { Width = 900, Height = 29, Margin = new Padding(0, 1, 0, 1) };
    var checkBox = new CheckBox { Text = label, AutoSize = true, Location = new Point(0, 2), Checked = false };
    panel.Controls.Add(checkBox);
    checks.Add(checkBox);
    var textBox = new TextBox { Location = new Point(430, 0), Width = 250, Text = defaultSuffix, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    suffixes.Add(textBox);
    panel.Controls.Add(textBox);
    // Extra "HF" (Hide Future dates) column - visible by default for applicable rows so the
    // user can toggle it directly; excluded rows (e.g. Previous period) are hidden further down.
    var hfCheck = new CheckBox { Text = "HF", AutoSize = true, Location = new Point(700, 2), Checked = false, Visible = true };
    panel.Controls.Add(hfCheck);
    hideFutureChecks.Add(hfCheck);
    // Extra "Cmp" (Comparison) column - only meaningful for YTD/QTD/MTD/RNm/RNm avg, hidden for
    // all other rows further down (see comparisonSupportedIndices).
    var cmpCheck = new CheckBox { Text = "Cmp", AutoSize = true, Location = new Point(755, 2), Checked = false, Visible = false };
    panel.Controls.Add(cmpCheck);
    comparisonChecks.Add(cmpCheck);
    descriptions.Add(description);
    categories.Add(category);
    calcPanels.Add(panel);
};

// ── Calculation checkboxes ─────────────────────────────────────────────────
// "*" marks calculations that need an N value - you'll be prompted for it in a follow-up dialog.
AddLabel("Choose Time Intelligence calculations and function name suffixes  (* = prompts for a value):");
AddLabel("\"HF\" = Hide Future Dates - excludes dates after the last available/historical date (configured in a follow-up dialog if any HF box is checked).");
AddLabel("\"Cmp\" = Comparison (only for YTD/QTD/WTD/MTD/RNm/RNm avg) - adds \"<suffix> Dev\" and \"<suffix> idx\" UDFs comparing the calc against its prior-period counterpart (PYTD/PQTD/PYWTD/PMTD/RNm LY/RNm avg LY) via Comparison.Deviation / Comparison.Pct_change.");

AddCalc("Accumulated", "Agg. year-to-date",                  "YTD",
        "Accumulated year to date (YTD)");                                       // 0
AddCalc("Accumulated", "Agg. year-to-date last year",         "PYTD",
        "Previous year-to-date (PYTD)");                                         // 1
AddCalc("Accumulated", "Agg. quarter-to-date",                "QTD",
        "Accumulated quarter to date (QTD)");                                    // 2
AddCalc("Accumulated", "Agg. quarter-to-date last year",      "PQTD",
        "Previous quarter-to-date (PQTD)");                                      // 3
AddCalc("Accumulated", "Agg. week-to-date",                   "WTD",
        "Accumulated week to date (WTD)");                                       // 4
AddCalc("Accumulated", "Agg. week-to-date last year",         "PYWTD",
        "Previous year week-to-date (PYWTD)");                                   // 5
AddCalc("Accumulated", "Agg. week-to-date previous week",     "PWTD",
        "Previous week-to-date, i.e. week-to-date shifted back one week (PWTD)"); // 6
AddCalc("Accumulated", "Agg. month-to-date",                  "MTD",
        "Accumulated month to date (MTD)");                                      // 7
AddCalc("Accumulated", "Agg. month-to-date last year",        "PMTD",
        "Previous month-to-date (PMTD)");                                        // 8
AddCalc("Accumulated", "Previous year (full year total)",     "PY",
        "Total for the entire previous year (accumulated), regardless of the current date selection (PY)"); // 9
AddCalc("Accumulated", "Previous month (full month total)",   "PM",
        "Total for the entire previous month (accumulated), regardless of the current date selection (PM)");  // 10
AddCalc("Moving total and averages", "Rolling N days *",                    "RNd",
        "Rolling period of N days, from reference date backwards");              // 11
AddCalc("Moving total and averages", "Rolling N months *",                  "RNm",
        "Rolling period of N months, from reference date backwards");            // 12
AddCalc("Moving total and averages", "Rolling N weeks *",                   "RNw",
        "Rolling period of N weeks, from reference date backwards");             // 13
AddCalc("Moving total and averages", "Rolling N months avg/month *",        "RNm avg",
        "Rolling N months average per month, from reference date backwards");    // 14
AddCalc("Moving total and averages", "Rolling N days avg/day *",            "RNd avg",
        "Rolling N days average per selling day (AVERAGEX excludes blank days), from reference date backwards. Reference: https://www.sqlbi.com/articles/the-hidden-complexity-of-a-daily-average-in-dax/"); // 15
AddCalc("Moving total and averages", "Running total",                       "RT",
        "Running total, from the first date available until selected date");     // 16
AddCalc("Date-filtering", "Previous period (standard Date column only)", "PP",
        "Previous period compared to the currently selected date range (PP)");   // 17
AddCalc("Accumulated", "Current year (full year total)",      "CY",
        "Total for the entire current year (accumulated), regardless of the current date selection (CY)"); // 18
AddCalc("Moving total and averages", "Rolling N months last year *", "RNm LY",
        "Rolling period of N months, from reference date backwards, shifted to the same period last year"); // 19
AddCalc("Moving total and averages", "Rolling N months avg/month last year *", "RNm avg LY",
        "Rolling N months average per month, from reference date backwards, shifted to the same period last year"); // 20
AddCalc("Moving total and averages", "Incoming balance", "IB",
        "Incoming balance - the accumulated total up to (but not including) the first selected/filtered date (IB)"); // 21
AddCalc("Semi-Additive", "LastDate", "LD",
        "Semi-additive balance - returns the value at the last date with data in the selected fact table, ignoring the current date filter except for the max date (LD). Reference: https://www.daxpatterns.com/semi-additive-calculations/"); // 22
AddCalc("Semi-Additive", "Open - End of period", "Open EOP",
        "Open at the end of the selected period - has the selected from date on/before the period end, and not yet the selected to date (or blank) (Open EOP). Reference: https://www.daxpatterns.com/semi-additive-calculations/"); // 23
AddCalc("Semi-Additive", "Total during period", "Open Total",
        "Open at any point during the selected period - has the selected from date on/before period end, and the selected to date after period start (or blank) (Open Total). Reference: https://www.daxpatterns.com/semi-additive-calculations/"); // 24
AddCalc("Date-filtering", "Simple filter - Historical dates", "Hist",
        "Returns the base measure only for historical dates (up to today), BLANK afterwards - pairs with a forecast measure for remaining year. Reference: https://www.sqlbi.com/articles/showing-actuals-and-forecasts-in-the-same-chart-with-power-bi/"); // 25
AddCalc("Date-filtering", "Simple filter - Remaining year", "Remain",
        "Returns the base measure only for future/remaining dates (after today), BLANK for historical dates - pairs with an actuals measure for the historical period. Reference: https://www.sqlbi.com/articles/showing-actuals-and-forecasts-in-the-same-chart-with-power-bi/"); // 26
AddCalc("Semi-Additive", "LastDate per Category", "LD Cat",
        "Semi-additive balance per category - returns the value at the last date with data in the selected fact table, broken down per category in the current filter context (e.g. per Product) (LD Cat). Reference: https://www.daxpatterns.com/semi-additive-calculations/"); // 27
AddCalc("Semi-Additive", "Previous per Category", "Prev Cat",
        "Semi-additive balance per category - returns the value at the date immediately before the last date with data in the selected fact table, broken down per category in the current filter context (Prev Cat). Reference: https://www.daxpatterns.com/semi-additive-calculations/"); // 28
// Appended after all other calls (rather than inserted alongside "Rolling N weeks") so the
// existing "checks[N]"/index-constant references above are not renumbered - see note below.
AddCalc("Moving total and averages", "Rolling N weeks avg/week *",          "RNw avg",
        "Rolling N weeks average per week, from reference date backwards");      // 29
// Appended to preserve all existing numeric indices used throughout the script.
AddCalc("Moving total and averages", "Rolling N days avg/transaction date *", "RNd avg observed",
        "Rolling N days average per observed transaction-date period. Includes calendar days without transactions inside the global first/last transaction-date boundaries, returns zero for an observed period without value, and blank outside the observed period. Reference: https://www.sqlbi.com/articles/the-hidden-complexity-of-a-daily-average-in-dax/"); // 30

const int RollingMonthsLyIndex       = 19;
const int RollingMonthsAvgLyIndex    = 20;
const int IncomingBalanceIndex       = 21;
const int LastDateIndex              = 22;
const int OpenEndOfPeriodIndex       = 23;
const int TotalDuringPeriodIndex     = 24;
const int HistoricalDatesIndex       = 25;
const int RemainingYearIndex         = 26;
const int LastDatePerCategoryIndex   = 27;
const int PreviousPerCategoryIndex   = 28;
const int RollingWeeksAvgIndex       = 29;
const int RollingDaysObservedAvgIndex = 30;

// Explicit display order override for the "Semi-Additive" category.
var semiAdditiveDisplayOrder = new List<int> {
    LastDateIndex, LastDatePerCategoryIndex, PreviousPerCategoryIndex,
    OpenEndOfPeriodIndex, TotalDuringPeriodIndex
};

// Explicit display order override for "Moving total and averages" only - places the appended
// "Rolling N weeks avg/week" (index 29) right after "Rolling N weeks" (index 13), instead of at
// the end where its insertion-order position would otherwise put it.
var movingDisplayOrder = new List<int> {
    11, 12, 13, RollingWeeksAvgIndex, 14, 15, RollingDaysObservedAvgIndex,
    RollingMonthsLyIndex, RollingMonthsAvgLyIndex, 16, IncomingBalanceIndex
};

foreach (var category in categoryOrder)
{
    AddSectionLabel(category + ":");
    if (category == "Semi-Additive")
    {
        foreach (var i in semiAdditiveDisplayOrder)
            flowPanel.Controls.Add(calcPanels[i]);
        continue;
    }
    if (category == "Moving total and averages")
    {
        foreach (var i in movingDisplayOrder)
            flowPanel.Controls.Add(calcPanels[i]);
        continue;
    }
    for (int i = 0; i < calcPanels.Count; i++)
        if (categories[i] == category)
            flowPanel.Controls.Add(calcPanels[i]);
}

// "Previous Period" only makes sense for a plain Date column (DATESBETWEEN on the raw selection),
// not for a named Calendar option. It is also already a fixed historical window, so it never gets
// a "hide future dates" (HF) variant.
hideFutureChecks[17].Visible = false;
hideFutureChecks[17].Enabled = false;
checks[17].Parent.Visible = !isCalendar;
if (isCalendar) checks[17].Checked = false;

// "Previous year", "Previous month", "Current year", LastDate, open-period, and the simple
// historical/remaining filters are fixed-period or non-time-accumulation patterns, so hiding
// future dates has no effect on them; treat them the same as "Previous Period" above.
hideFutureChecks[9].Visible  = false;
hideFutureChecks[9].Enabled  = false;
hideFutureChecks[10].Visible = false;
hideFutureChecks[10].Enabled = false;
hideFutureChecks[18].Visible = false;
hideFutureChecks[18].Enabled = false;
hideFutureChecks[LastDateIndex].Visible = false;
hideFutureChecks[LastDateIndex].Enabled = false;
hideFutureChecks[OpenEndOfPeriodIndex].Visible = false;
hideFutureChecks[OpenEndOfPeriodIndex].Enabled = false;
hideFutureChecks[TotalDuringPeriodIndex].Visible = false;
hideFutureChecks[TotalDuringPeriodIndex].Enabled = false;
hideFutureChecks[HistoricalDatesIndex].Visible = false;
hideFutureChecks[HistoricalDatesIndex].Enabled = false;
hideFutureChecks[RemainingYearIndex].Visible = false;
hideFutureChecks[RemainingYearIndex].Enabled = false;
hideFutureChecks[LastDatePerCategoryIndex].Visible = false;
hideFutureChecks[LastDatePerCategoryIndex].Enabled = false;
hideFutureChecks[PreviousPerCategoryIndex].Visible = false;
hideFutureChecks[PreviousPerCategoryIndex].Enabled = false;
hideFutureChecks[RollingDaysObservedAvgIndex].Visible = false;
hideFutureChecks[RollingDaysObservedAvgIndex].Enabled = false;

var hideFutureSupportedIndices = new HashSet<int>
{
    0, 1, 2, 3, 4, 5, 6, 7, 8,
    11, 12, 13, 14, 15, 16,
    RollingMonthsLyIndex, RollingMonthsAvgLyIndex, IncomingBalanceIndex, RollingWeeksAvgIndex
};

// Wire per-row "Hide Future" (HF) checkboxes: toggling updates the suffix textbox itself -
// checked appends " HF" to the suffix; unchecked reverts to the plain SQLBI-convention suffix
// (the "complete", un-filtered calculation). Off by default for every row - the user opts in
// per calc. See https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
for (int i = 0; i < checks.Count; i++)
{
    if (!hideFutureSupportedIndices.Contains(i)) continue;

    string baseSuffix = suffixes[i].Text;
    var hfCheck       = hideFutureChecks[i];
    var suffixBox     = suffixes[i];
    hfCheck.CheckedChanged += (s, e) => {
        suffixBox.Text = hfCheck.Checked ? baseSuffix + " HF" : baseSuffix;
    };
}

// "Comparison" (Cmp) column - only applicable to YTD, QTD, WTD, MTD, RNm and RNm avg, which each
// have a natural prior-period counterpart (PYTD, PQTD, PYWTD, PMTD, RNm LY, RNm avg LY). When
// checked, two extra UDFs get created later ("<suffix> Dev" and "<suffix> idx") comparing the
// calc against that counterpart via the model-independent Comparison.Deviation / Comparison.Pct_change UDFs.
var comparisonSupportedIndices = new HashSet<int> { 0, 2, 4, 7, 12, 14 };
for (int i = 0; i < comparisonChecks.Count; i++)
    comparisonChecks[i].Visible = comparisonSupportedIndices.Contains(i);

// Flags calculations that only use "classic" date logic (MAX/MIN/YEAR/DATESBETWEEN-style raw
// date comparisons rather than calendar time-unit DAX functions like DATESYTD/SAMEPERIODLASTYEAR)
// - these behave identically whether a Calendar option or a plain Date column was picked, so the
// Calendar-specific settings (fiscal year end, custom week numbering, etc.) don't apply to them.
// Only worth flagging when a Calendar option was actually selected as the date source.
var classicDateOnlyIndices = new HashSet<int> {
    9, 10, 16, 17, 18,
    IncomingBalanceIndex, LastDateIndex, OpenEndOfPeriodIndex, TotalDuringPeriodIndex,
    HistoricalDatesIndex, RemainingYearIndex, LastDatePerCategoryIndex, PreviousPerCategoryIndex,
    RollingDaysObservedAvgIndex
};
if (isCalendar)
{
    var classicDateTip = new ToolTip();
    foreach (var i in classicDateOnlyIndices)
    {
        checks[i].Text += "  (classic date)";
        classicDateTip.SetToolTip(checks[i],
            "Uses plain MAX/MIN/date-comparison logic, not a calendar time-unit function - " +
            "behaves the same regardless of Calendar-specific options (fiscal year end, custom " +
            "week numbering, etc.).");
    }
}

// Show/hide rows based on the Date source picked in the previous dialog - fixed for this run,
// so no dynamic re-evaluation is needed (the date source can no longer change at this point).
checks[4].Parent.Visible  = isWeekly;
checks[5].Parent.Visible  = isWeekly;
checks[6].Parent.Visible  = isWeekly;
checks[13].Parent.Visible = isWeekly;
checks[RollingWeeksAvgIndex].Parent.Visible = isWeekly;

checks[2].Parent.Visible = hasQuarter;
checks[3].Parent.Visible = hasQuarter;

checks[7].Parent.Visible = hasMonth;
checks[8].Parent.Visible = hasMonth;

bool showMonthRolling = !isWeekly;
checks[12].Parent.Visible = showMonthRolling;
checks[14].Parent.Visible = showMonthRolling;
checks[15].Parent.Visible = showMonthRolling;
checks[RollingMonthsLyIndex].Parent.Visible = showMonthRolling;
checks[RollingMonthsAvgLyIndex].Parent.Visible = showMonthRolling;

// ── Fiscal Year option (applies only to YTD, since DATESYTD is the only time-intel
//    function that accepts a fiscal year-end date) ────────────────────────────────
AddLabel("Fiscal Year:");
var fiscalYearPanel    = new Panel { AutoSize = true };
var fiscalYearCheckbox = new CheckBox
{
    Text     = "Also create Fiscal Year variants of YTD (TI_YTD_FY / TI_PYTD_FY)",
    AutoSize = true,
    Location = new Point(0, 2),
    Checked  = false
};
fiscalYearPanel.Controls.Add(fiscalYearCheckbox);
flowPanel.Controls.Add(fiscalYearPanel);

// Note: detailed "Hide future dates" configuration, the Rolling-N values, the "Rolling N
// months avg/month" Year-Month column choice, and the semi-additive fact-date pickers are NOT
// asked here inline - they're only prompted via small follow-up dialogs after clicking OK below,
// and only if actually needed.

// ── Name / folder options ───────────────────────────────────────────────────
AddLabel("(Optional) Function name prefix  [default: Local.TimeIntelligence]:");
var prefixTextbox = new TextBox { Width = 350, Text = "Local.TimeIntelligence" };
flowPanel.Controls.Add(prefixTextbox);

AddLabel("(Optional) FolderSuffix annotation value  [default: Time Intelligence]:");
var folderSuffixTextbox = new TextBox { Width = 350, Text = "Time Intelligence" };
flowPanel.Controls.Add(folderSuffixTextbox);

// ── Buttons ──────────────────────────────────────────────────────────────────
var buttonPanel  = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 20, 0, 0) };
var okButton     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
buttonPanel.Controls.Add(okButton);
buttonPanel.Controls.Add(cancelButton);
flowPanel.Controls.Add(buttonPanel);

aggForm.AcceptButton = okButton;
aggForm.CancelButton = cancelButton;

Action resizeCalculationRows = () => {
    int rowWidth = Math.Max(760, scrollPanel.ClientSize.Width - 35);
    flowPanel.MinimumSize = new Size(rowWidth, 0);
    for (int i = 0; i < calcPanels.Count; i++)
    {
        calcPanels[i].Width = rowWidth;
        int hfX = rowWidth - 155;
        int cmpX = rowWidth - 100;
        suffixes[i].Width = Math.Max(170, hfX - suffixes[i].Left - 20);
        hideFutureChecks[i].Left = hfX;
        comparisonChecks[i].Left = cmpX;
    }
};
scrollPanel.ClientSizeChanged += (s, e) => resizeCalculationRows();
aggForm.Shown += (s, e) => resizeCalculationRows();

if (aggForm.ShowDialog() != DialogResult.OK) return;

// ===================== DIALOG 3: Rolling N values (only asked if needed) =====================
bool needDays   = checks[11].Checked || checks[15].Checked || checks[RollingDaysObservedAvgIndex].Checked;
bool needMonths = checks[12].Checked || checks[14].Checked || checks[RollingMonthsLyIndex].Checked || checks[RollingMonthsAvgLyIndex].Checked;
bool needWeeks  = checks[13].Checked || checks[RollingWeeksAvgIndex].Checked;

int nDays = 365, nMonths = 12, nWeeks = 4;
var suffixPreviewMap = new List<(int idx, TextBox box)>();

if (needDays || needMonths || needWeeks)
{
    var rollForm = new Form
    {
        Text = "Rolling period values",
        MaximizeBox = true,
        MinimizeBox = false,
        ShowIcon = false,
        Size = new Size(
            Math.Min(760, Screen.PrimaryScreen.WorkingArea.Width - 100),
            Math.Min(720, Screen.PrimaryScreen.WorkingArea.Height - 100)),
        MinimumSize = new Size(520, 360),
        FormBorderStyle = FormBorderStyle.Sizable,
        StartPosition = FormStartPosition.CenterScreen
    };
    var rollFlow = new FlowLayoutPanel {
        Dock = DockStyle.Fill, Padding = new Padding(12), FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true
    };
    rollForm.Controls.Add(rollFlow);
    Action<string> rollLabel = s => rollFlow.Controls.Add(new Label { Margin = new Padding(0,10,0,5), Text = s, AutoSize = true });

    NumericUpDown rollingDaysInput = null, rollingMonthsInput = null, rollingWeeksInput = null;

    // Adds an editable "resulting suffix" preview box for a checked row, tied to the given N
    // spinner - it re-generates the suffix (replacing "N") whenever the spinner value changes,
    // but the user can still freely retype it afterwards before clicking OK.
    Action<int, NumericUpDown> AddSuffixPreview = (idx, input) => {
        if (!checks[idx].Checked) return;
        string orig = suffixes[idx].Text;
        var box = new TextBox { Width = 220, Text = orig.Replace("N", ((int)input.Value).ToString()) };
        rollFlow.Controls.Add(box);
        input.ValueChanged += (s, e) => box.Text = orig.Replace("N", ((int)input.Value).ToString());
        suffixPreviewMap.Add((idx, box));
    };

    if (needDays)
    {
        rollLabel("Rolling N days  (default 365 — used for Rolling N days and both daily-average variants):");
        rollingDaysInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 3650, Value = 365 };
        rollFlow.Controls.Add(rollingDaysInput);
        rollLabel("Resulting suffix(es) (editable):");
        AddSuffixPreview(11, rollingDaysInput);
        AddSuffixPreview(15, rollingDaysInput);
        AddSuffixPreview(RollingDaysObservedAvgIndex, rollingDaysInput);
    }
    if (needMonths)
    {
        rollLabel("Rolling N months  (default 12 — used for Rolling N months, Rolling N months avg/month, and last-year rolling-month variants):");
        rollingMonthsInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 120, Value = 12 };
        rollFlow.Controls.Add(rollingMonthsInput);
        rollLabel("Resulting suffix(es) (editable):");
        AddSuffixPreview(12, rollingMonthsInput);
        AddSuffixPreview(14, rollingMonthsInput);
        AddSuffixPreview(RollingMonthsLyIndex, rollingMonthsInput);
        AddSuffixPreview(RollingMonthsAvgLyIndex, rollingMonthsInput);
    }
    if (needWeeks)
    {
        rollLabel("Rolling N weeks  (used for Rolling N weeks and Rolling N weeks avg/week):");
        rollingWeeksInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 520, Value = 4 };
        rollFlow.Controls.Add(rollingWeeksInput);
        rollLabel("Resulting suffix(es) (editable):");
        AddSuffixPreview(13, rollingWeeksInput);
        AddSuffixPreview(RollingWeeksAvgIndex, rollingWeeksInput);
    }

    var rollBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
    var rollOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
    var rollCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    rollBtnPanel.Controls.Add(rollOk);
    rollBtnPanel.Controls.Add(rollCancel);
    rollFlow.Controls.Add(rollBtnPanel);
    rollForm.AcceptButton = rollOk;
    rollForm.CancelButton = rollCancel;

    if (rollForm.ShowDialog() != DialogResult.OK) return;

    if (rollingDaysInput   != null) nDays   = (int)rollingDaysInput.Value;
    if (rollingMonthsInput != null) nMonths = (int)rollingMonthsInput.Value;
    if (rollingWeeksInput  != null) nWeeks  = (int)rollingWeeksInput.Value;
}

// ===================== Resolve remaining selections =====================
var namePrefix   = string.IsNullOrWhiteSpace(prefixTextbox.Text)
                     ? "Local.TimeIntelligence"
                     : prefixTextbox.Text.Trim();

var folderSuffix = folderSuffixTextbox.Text.Trim();

Func<int, int> NFor = idx => (idx == 11 || idx == 15 || idx == RollingDaysObservedAvgIndex)
    ? nDays
    : (idx == 13 || idx == RollingWeeksAvgIndex)
        ? nWeeks
        : nMonths;

// Apply the (possibly user-edited) final suffix from the Rolling N values dialog, and replace
// the literal "N" in each applicable row's description.
foreach (var (idx, box) in suffixPreviewMap)
{
    suffixes[idx].Text = box.Text;
    descriptions[idx]  = descriptions[idx].Replace("N", NFor(idx).ToString());
}

Func<string, string> FuncName = suffix =>
    namePrefix + ".TI_" + suffix.Replace(" ", "_");

var skipped = new List<string>();
Func<string, string> NormalizeExpr = s => string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));

Action<string, string, string, string> CreateUDF = (funcName, daxBody, measureSuffix, description) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    // ---- Check 1: an identical DAX expression already exists under a different function name ----
    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        skipped.Add($"'{funcName}' was NOT created - using the existing expression '{dupExprFunc.Name}' instead, since it is identical.");
        return;
    }

    // ---- Check 2: a function with this exact name already exists ----
    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            skipped.Add($"'{funcName}' already exists with the same expression - using the existing definition, nothing changed.");
            return;
        }

        // Custom 3-way dialog (instead of a plain rename prompt) so a name conflict doesn't keep
        // re-appearing on every run: the user can Rename, Overwrite the existing (outdated)
        // definition, or Skip.
        var conflictDlg = new Form {
            Text = "Function name already exists", Width = 680, Height = 340,
            MinimumSize = new Size(520, 300), FormBorderStyle = FormBorderStyle.Sizable,
            StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false
        };
        var conflictFlow = new FlowLayoutPanel {
            Dock = DockStyle.Fill, Padding = new Padding(12), FlowDirection = FlowDirection.TopDown,
            WrapContents = false, AutoScroll = true
        };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = $"A function named '{funcName}' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(440, 0), Margin = new Padding(0, 0, 0, 10) });

        var renameRow    = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var renameRadio  = new RadioButton { Text = "Rename to:", AutoSize = true, Checked = true, Margin = new Padding(0, 3, 8, 0) };
        var renameBox    = new TextBox { Width = 250, Text = funcName };
        renameRow.Controls.Add(renameRadio);
        renameRow.Controls.Add(renameBox);
        conflictFlow.Controls.Add(renameRow);

        var overwriteRadio = new RadioButton { Text = "Overwrite the existing definition with this new one", AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var skipRadio       = new RadioButton { Text = "Skip - do not create this UDF", AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        conflictFlow.Controls.Add(overwriteRadio);
        conflictFlow.Controls.Add(skipRadio);

        var conflictBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
        var conflictOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
        var conflictCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        conflictBtnPanel.Controls.Add(conflictOk);
        conflictBtnPanel.Controls.Add(conflictCancel);
        conflictFlow.Controls.Add(conflictBtnPanel);
        conflictDlg.AcceptButton = conflictOk;
        conflictDlg.CancelButton = conflictCancel;

        if (conflictDlg.ShowDialog() != DialogResult.OK)
        {
            skipped.Add($"'{funcName}' was NOT created - conflict dialog cancelled.");
            return;
        }

        if (overwriteRadio.Checked)
        {
            existing.Expression  = daxBody;
            existing.Description = description;
            existing.SetAnnotation("FunctionCategory", "Time Intelligence");
            existing.SetAnnotation("MeasureSuffix",    measureSuffix);
            existing.SetAnnotation("FolderSuffix",     folderSuffix);
            skipped.Add($"'{funcName}' already existed with a different expression - its definition was OVERWRITTEN with the new one.");
            return;
        }
        else if (skipRadio.Checked)
        {
            skipped.Add($"'{funcName}' was NOT created - a function with that name already exists (existing definition kept).");
            return;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                skipped.Add($"'{funcName}' was NOT created - no valid new name was provided (existing definition kept).");
                return;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                skipped.Add($"'{funcName}' was NOT created - the alternate name '{newName}' also already exists.");
                return;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
    fn.SetAnnotation("FunctionCategory", "Time Intelligence");
    fn.SetAnnotation("MeasureSuffix",    measureSuffix);
    fn.SetAnnotation("FolderSuffix",     folderSuffix);
};

// ===================== Follow-up dialog: Year-Month column =====================
// Only asked if a rolling-month average calculation was actually checked. If the selected
// Calendar option already has a column mapped to TimeUnit.Month (the standard, year-unique
// "Year-Month" granularity - e.g. "Year month no"), that mapping is used automatically and no
// dialog is shown. Otherwise (plain Date column, or a Calendar option without a Month mapping),
// the user is prompted to pick a suitable Int64 column.
var yearMonthColumn = "Calendar[Year month no]";
if (checks[14].Checked || checks[RollingMonthsAvgLyIndex].Checked)
{
    var monthTimeUnit = isCalendar ? cal.GetTimeUnits().FirstOrDefault(tu => tu.TimeUnit == TimeUnit.Month) : null;

    if (monthTimeUnit?.PrimaryColumn != null)
    {
        yearMonthColumn = monthTimeUnit.PrimaryColumn.DaxObjectFullName;
    }
    else
    {
        var monthYearColumns = Model.AllColumns
            .Where(c => !IsDisconnected(c.Table))
            .Where(c => c.DataType == DataType.Int64 &&
                        (c.Name.IndexOf("month", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         c.Name.IndexOf("year", StringComparison.OrdinalIgnoreCase) >= 0))
            .ToList();
        var intColumns = monthYearColumns.Count > 0
            ? monthYearColumns
            : Model.AllColumns.Where(c => !IsDisconnected(c.Table) && c.DataType == DataType.Int64).ToList();

        var ymDlg = new Form { Text = "Rolling N months avg/month", Width = 700, Height = 230, MinimumSize = new Size(520, 220), FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false };
        var ymLbl = new Label { Text = "Choose a unique Year-Month column (used to average rolling months):", AutoSize = true, Location = new Point(12, 12), MaximumSize = new Size(650, 0) };
        var ymCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(12, 48), Width = 650, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        intColumns.ForEach(c => ymCombo.Items.Add(new ComboBoxItem(c)));
        if (ymCombo.Items.Count > 0) ymCombo.SelectedIndex = 0;
        var ymBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Location = new Point(490, 105), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var ymOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
        var ymCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        ymBtnPanel.Controls.Add(ymOk);
        ymBtnPanel.Controls.Add(ymCancel);
        ymDlg.Controls.Add(ymLbl);
        ymDlg.Controls.Add(ymCombo);
        ymDlg.Controls.Add(ymBtnPanel);
        ymDlg.AcceptButton = ymOk;
        ymDlg.CancelButton = ymCancel;

        if (ymDlg.ShowDialog() == DialogResult.OK && ymCombo.SelectedItem != null)
            yearMonthColumn = (ymCombo.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;
    }
}

// ===================== Follow-up dialog: Year-Week column =====================
// Mirrors the Year-Month resolution above, for "Rolling N weeks avg/week". If the selected
// Calendar option has a column mapped to TimeUnit.Week (the year-unique "Year-Week" granularity -
// e.g. "ISO year week"), it is used automatically. Otherwise the user is prompted; week columns
// are commonly text (e.g. "2024-W01"), so both Int64 and String columns are offered.
var yearWeekColumn = "Calendar[ISO year week]";
if (checks[RollingWeeksAvgIndex].Checked)
{
    var weekTimeUnit = isCalendar ? cal.GetTimeUnits().FirstOrDefault(tu => tu.TimeUnit == TimeUnit.Week) : null;

    if (weekTimeUnit?.PrimaryColumn != null)
    {
        yearWeekColumn = weekTimeUnit.PrimaryColumn.DaxObjectFullName;
    }
    else
    {
        var weekNameColumns = Model.AllColumns
            .Where(c => !IsDisconnected(c.Table))
            .Where(c => (c.DataType == DataType.Int64 || c.DataType == DataType.String) &&
                        c.Name.IndexOf("week", StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();
        var weekColumns = weekNameColumns.Count > 0
            ? weekNameColumns
            : Model.AllColumns.Where(c => !IsDisconnected(c.Table) &&
                                          (c.DataType == DataType.Int64 || c.DataType == DataType.String)).ToList();

        var ywDlg = new Form { Text = "Rolling N weeks avg/week", Width = 700, Height = 230, MinimumSize = new Size(520, 220), FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false };
        var ywLbl = new Label { Text = "Choose a unique Year-Week column (used to average rolling weeks):", AutoSize = true, Location = new Point(12, 12), MaximumSize = new Size(650, 0) };
        var ywCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(12, 48), Width = 650, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        weekColumns.ForEach(c => ywCombo.Items.Add(new ComboBoxItem(c)));
        if (ywCombo.Items.Count > 0) ywCombo.SelectedIndex = 0;
        var ywBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Location = new Point(490, 105), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var ywOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
        var ywCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        ywBtnPanel.Controls.Add(ywOk);
        ywBtnPanel.Controls.Add(ywCancel);
        ywDlg.Controls.Add(ywLbl);
        ywDlg.Controls.Add(ywCombo);
        ywDlg.Controls.Add(ywBtnPanel);
        ywDlg.AcceptButton = ywOk;
        ywDlg.CancelButton = ywCancel;

        if (ywDlg.ShowDialog() == DialogResult.OK && ywCombo.SelectedItem != null)
            yearWeekColumn = (ywCombo.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;
    }
}

// ===================== Follow-up dialog: LastDate fact date column =====================
string lastDateFactColumn = null;
string lastDateFactTableRef = null;

if (checks[LastDateIndex].Checked || checks[LastDatePerCategoryIndex].Checked || checks[PreviousPerCategoryIndex].Checked)
{
    var lastDateColumns = Model.AllColumns
        .Where(c => c.DataType == DataType.DateTime &&
                    c.Table != dateTable &&
                    !IsDisconnected(c.Table) &&
                    Model.Relationships.Any(r => r.IsActive && r.FromTable == c.Table && r.FromCardinality == RelationshipEndCardinality.Many))
        .ToList();

    if (lastDateColumns.Count == 0)
    {
        Info("No eligible fact-table DateTime columns were found for LastDate / LastDate per Category / Previous per Category. Those UDFs will be skipped.");
    }
    else
    {
        var lastDateDlg = new Form { Text = "LastDate", Width = 700, Height = 240, MinimumSize = new Size(520, 220), FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false };
        var lastDateLbl = new Label { Text = "Choose the fact-table date column to use for LastDate / LastDate per Category / Previous per Category:", AutoSize = true, Location = new Point(12, 12), MaximumSize = new Size(650, 0) };
        var lastDateCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(12, 60), Width = 650, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        lastDateColumns.ForEach(c => lastDateCombo.Items.Add(new ComboBoxItem(c)));
        if (lastDateCombo.Items.Count > 0) lastDateCombo.SelectedIndex = 0;
        var lastDateBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Location = new Point(490, 115), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var lastDateOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
        var lastDateCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        lastDateBtnPanel.Controls.Add(lastDateOk);
        lastDateBtnPanel.Controls.Add(lastDateCancel);
        lastDateDlg.Controls.Add(lastDateLbl);
        lastDateDlg.Controls.Add(lastDateCombo);
        lastDateDlg.Controls.Add(lastDateBtnPanel);
        lastDateDlg.AcceptButton = lastDateOk;
        lastDateDlg.CancelButton = lastDateCancel;

        if (lastDateDlg.ShowDialog() == DialogResult.OK && lastDateCombo.SelectedItem != null)
        {
            var selectedFactDate = (lastDateCombo.SelectedItem as ComboBoxItem).Column;
            lastDateFactColumn = selectedFactDate.DaxObjectFullName;
            lastDateFactTableRef = $"'{selectedFactDate.Table.Name.Replace("'", "''")}'";
            descriptions[LastDateIndex] = $"Semi-additive balance - returns the value at the last date with data in {lastDateFactColumn}, ignoring the current date filter except for the max date (LD). Reference: https://www.daxpatterns.com/semi-additive-calculations/";
            descriptions[LastDatePerCategoryIndex] = $"Semi-additive balance per category - returns the value at the last date with data in {lastDateFactColumn}, broken down per category in the current filter context (LD Cat). Reference: https://www.daxpatterns.com/semi-additive-calculations/";
            descriptions[PreviousPerCategoryIndex] = $"Semi-additive balance per category - returns the value at the date immediately before the last date with data in {lastDateFactColumn}, broken down per category in the current filter context (Prev Cat). Reference: https://www.daxpatterns.com/semi-additive-calculations/";
        }
    }
}

// ===================== Follow-up dialog: observed transaction-date column =====================
// Used by "Rolling N days avg/transaction date". The UDF is intentionally local/model-dependent:
// it derives FirstObservedDate and LastObservedDate from this fact-table column after REMOVEFILTERS().
string observedTransactionDateColumn = null;

if (checks[RollingDaysObservedAvgIndex].Checked)
{
    var observedDateColumns = Model.AllColumns
        .Where(c => c.DataType == DataType.DateTime &&
                    c.Table != dateTable &&
                    !IsDisconnected(c.Table) &&
                    Model.Relationships.Any(r =>
                        r.IsActive &&
                        r.FromTable == c.Table &&
                        r.FromCardinality == RelationshipEndCardinality.Many))
        .OrderBy(c => c.Table.Name)
        .ThenBy(c => c.Name)
        .ToList();

    if (observedDateColumns.Count == 0)
    {
        Info("No eligible fact-table DateTime columns were found for Rolling N days avg/transaction date. That UDF will be skipped.");
    }
    else
    {
        var observedDlg = new Form
        {
            Text = "Rolling N days avg/transaction date",
            Width = 700,
            Height = 230,
            MinimumSize = new Size(560, 210),
            FormBorderStyle = FormBorderStyle.Sizable,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = true,
            MinimizeBox = false
        };
        var observedLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            RowCount = 4,
            ColumnCount = 1
        };
        observedLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        observedLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        observedLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        observedLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var observedLbl = new Label
        {
            Text = "Choose the fact-table transaction date used to derive the global FirstObservedDate and LastObservedDate boundaries:",
            AutoSize = true,
            MaximumSize = new Size(640, 0),
            Dock = DockStyle.Fill
        };
        var observedReference = new LinkLabel
        {
            Text = "Reference: SQLBI - The hidden complexity of a daily average in DAX",
            AutoSize = true,
            Dock = DockStyle.Fill
        };
        observedReference.LinkClicked += (s, e) =>
            System.Diagnostics.Process.Start("https://www.sqlbi.com/articles/the-hidden-complexity-of-a-daily-average-in-dax/");
        var observedCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Top,
            Width = 640
        };
        observedDateColumns.ForEach(c => observedCombo.Items.Add(new ComboBoxItem(c)));
        observedCombo.SelectedIndex = 0;
        var observedButtons = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            Dock = DockStyle.Fill
        };
        var observedOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
        var observedCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        observedButtons.Controls.Add(observedOk);
        observedButtons.Controls.Add(observedCancel);
        observedLayout.Controls.Add(observedLbl, 0, 0);
        observedLayout.Controls.Add(observedReference, 0, 1);
        observedLayout.Controls.Add(observedCombo, 0, 2);
        observedLayout.Controls.Add(observedButtons, 0, 3);
        observedDlg.Controls.Add(observedLayout);
        observedDlg.AcceptButton = observedOk;
        observedDlg.CancelButton = observedCancel;

        if (observedDlg.ShowDialog() != DialogResult.OK) return;

        observedTransactionDateColumn =
            (observedCombo.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;
        descriptions[RollingDaysObservedAvgIndex] =
            $"Rolling {nDays} days average per observed transaction-date period. Observation boundaries are derived from {observedTransactionDateColumn} with REMOVEFILTERS(); calendar days without transactions inside those boundaries are included, an observed period without value returns zero, and a period outside the boundaries returns blank. Reference: https://www.sqlbi.com/articles/the-hidden-complexity-of-a-daily-average-in-dax/";
    }
}

// ===================== Follow-up dialog: Open / total during period date columns =====================
string fromDateCol = null;
string toDateCol   = null;

if (checks[OpenEndOfPeriodIndex].Checked || checks[TotalDuringPeriodIndex].Checked)
{
    if (factDateColumns.Count == 0)
    {
        Info("No eligible fact-table DateTime columns were found for the open-period calculations. Those UDFs will be skipped.");
    }
    else
    {
        var openDlg = new Form { Text = "Open-period date columns", Width = 720, Height = 310, MinimumSize = new Size(540, 280), FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false };
        var fromLbl = new Label { Text = "From date column (e.g. Start Date):", AutoSize = true, Location = new Point(12, 12), MaximumSize = new Size(670, 0) };
        var fromCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(12, 42), Width = 670, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        factDateColumns.ForEach(c => fromCombo.Items.Add(new ComboBoxItem(c)));
        if (fromCombo.Items.Count > 0) fromCombo.SelectedIndex = 0;
        var toLbl = new Label { Text = "To date column (e.g. End Date):", AutoSize = true, Location = new Point(12, 88), MaximumSize = new Size(670, 0) };
        var toCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(12, 118), Width = 670, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        factDateColumns.ForEach(c => toCombo.Items.Add(new ComboBoxItem(c)));
        if (toCombo.Items.Count > 0) toCombo.SelectedIndex = 0;
        var openBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Location = new Point(510, 175), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var openOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
        var openCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        openBtnPanel.Controls.Add(openOk);
        openBtnPanel.Controls.Add(openCancel);
        openDlg.Controls.Add(fromLbl);
        openDlg.Controls.Add(fromCombo);
        openDlg.Controls.Add(toLbl);
        openDlg.Controls.Add(toCombo);
        openDlg.Controls.Add(openBtnPanel);
        openDlg.AcceptButton = openOk;
        openDlg.CancelButton = openCancel;

        if (openDlg.ShowDialog() == DialogResult.OK && fromCombo.SelectedItem != null && toCombo.SelectedItem != null)
        {
            fromDateCol = (fromCombo.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;
            toDateCol   = (toCombo.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;
            descriptions[OpenEndOfPeriodIndex] = $"Open at the end of the selected period - has {fromDateCol} on/before the period end, and not yet {toDateCol} (or blank) (Open EOP). Reference: https://www.daxpatterns.com/semi-additive-calculations/";
            descriptions[TotalDuringPeriodIndex] = $"Open at any point during the selected period - has {fromDateCol} on/before period end, and {toDateCol} after period start (or blank) (Open Total). Reference: https://www.daxpatterns.com/semi-additive-calculations/";
        }
    }
}

// ===================== Follow-up dialog: Hide future dates (HF) =====================
// Only asked if at least one applicable calculation has its "HF" checkbox checked.
// See https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
// Combines both the "historical dates" test and the CALCULATE-wrapping into a SINGLE DAX UDF,
// under the same namespace as the time intelligence functions with an extra ".Filtering." segment
// (e.g. Local.TimeIntelligence.Filtering.HideFutureDates), annotated with MeasureSuffix "HF".
string hfFuncFullName = null;
bool anyHideFutureRequested = hideFutureSupportedIndices.Any(i => checks[i].Checked && hideFutureChecks[i].Checked);

if (anyHideFutureRequested)
{
    var existingFuncs = Model.Functions.Select(f => f.Name).OrderBy(n => n).ToList();

    var hfDlg   = new Form { Text = "Configure \"Hide future dates\" (HF)", Width = 760, Height = 560, MinimumSize = new Size(620, 460), FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = true, MinimizeBox = false };
    var hfFlow  = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(12), AutoScroll = true };
    hfDlg.Controls.Add(hfFlow);

    Action<string> hfLabel = s => hfFlow.Controls.Add(new Label { Text = s, AutoSize = true, Margin = new Padding(0, 8, 0, 3) });

    hfLabel("Historical-dates filter function to use (pick an existing one, or create a new one):");
    var funcCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 680 };
    funcCombo.Items.Add("<Create new...>");
    foreach (var n in existingFuncs) funcCombo.Items.Add(n);
    funcCombo.SelectedIndex = 0;
    hfFlow.Controls.Add(funcCombo);

    var methodPanel = new Panel { AutoSize = true, Height = 70 };
    var methodA = new RadioButton { Text = "a) A column in the date table marks historical rows",  AutoSize = true, Location = new Point(0, 0),  Checked = true };
    var methodB = new RadioButton { Text = "b) Last date with transactions in a fact table",        AutoSize = true, Location = new Point(0, 23) };
    var methodC = new RadioButton { Text = "c) Last selected date, but not later than TODAY()",     AutoSize = true, Location = new Point(0, 46) };
    methodPanel.Controls.Add(methodA);
    methodPanel.Controls.Add(methodB);
    methodPanel.Controls.Add(methodC);
    hfFlow.Controls.Add(methodPanel);

    hfLabel("(a) Historical flag column (from the date table):");
    var colLabel = hfFlow.Controls[hfFlow.Controls.Count - 1] as Label;
    var colCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 680 };
    if (dateTable != null) foreach (var c in dateTable.Columns) colCombo.Items.Add(new ComboBoxItem(c));
    if (colCombo.Items.Count > 0) colCombo.SelectedIndex = 0;
    hfFlow.Controls.Add(colCombo);

    hfLabel("(a) That column represents \"historical\" as:");
    var valLabel = hfFlow.Controls[hfFlow.Controls.Count - 1] as Label;
    var valCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 250 };
    valCombo.Items.Add("1 / 0 (numeric)");
    valCombo.Items.Add("TRUE / FALSE (boolean)");
    valCombo.Items.Add("Yes / No (text)");
    valCombo.SelectedIndex = 0;
    hfFlow.Controls.Add(valCombo);

    hfLabel("(b) Transaction date column (any DateTime column, from a fact table other than the date table):");
    var factColLabel = hfFlow.Controls[hfFlow.Controls.Count - 1] as Label;
    var factColCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 680 };
    factDateColumns.ForEach(c => factColCombo.Items.Add(new ComboBoxItem(c)));
    if (factColCombo.Items.Count > 0) factColCombo.SelectedIndex = 0;
    hfFlow.Controls.Add(factColCombo);
    hfFlow.ClientSizeChanged += (s, e) => {
        int inputWidth = Math.Max(480, hfFlow.ClientSize.Width - 35);
        funcCombo.Width = inputWidth;
        colCombo.Width = inputWidth;
        factColCombo.Width = inputWidth;
    };

    Action updateHfVis = () => {
        bool createNew = funcCombo.SelectedIndex == 0;
        methodPanel.Visible = createNew;
        bool a = createNew && methodA.Checked;
        bool b = createNew && methodB.Checked;
        colLabel.Visible = a; colCombo.Visible = a; valLabel.Visible = a; valCombo.Visible = a;
        factColLabel.Visible = b; factColCombo.Visible = b;
    };
    funcCombo.SelectedIndexChanged += (s, e) => updateHfVis();
    methodA.CheckedChanged += (s, e) => updateHfVis();
    methodB.CheckedChanged += (s, e) => updateHfVis();
    methodC.CheckedChanged += (s, e) => updateHfVis();
    updateHfVis();

    var hfBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
    var hfOk       = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
    var hfCancel   = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    hfBtnPanel.Controls.Add(hfOk);
    hfBtnPanel.Controls.Add(hfCancel);
    hfFlow.Controls.Add(hfBtnPanel);
    hfDlg.AcceptButton = hfOk;
    hfDlg.CancelButton = hfCancel;

    if (hfDlg.ShowDialog() == DialogResult.OK)
    {
        if (funcCombo.SelectedIndex > 0)
        {
            hfFuncFullName = funcCombo.SelectedItem.ToString();
        }
        else
        {
            string body = null, desc = null;
            string hfFuncName = namePrefix + ".Filtering.HideFutureDates";

            if (methodA.Checked)
            {
                var col = (colCombo.SelectedItem as ComboBoxItem)?.Column;
                if (col == null)
                {
                    Info("No historical flag column selected. Skipping 'Hide future dates'.");
                }
                else
                {
                    string literal = valCombo.SelectedIndex switch { 1 => "TRUE()", 2 => "\"Yes\"", _ => "1" };
                    body = $@"(
    // Expression to evaluate for the historical dates only; its result is returned by this function
    valueExpr: ANYREF EXPR
) =>
    // https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    CALCULATE(
        valueExpr,
        KEEPFILTERS( {col.DaxObjectFullName} = {literal} )
    )";
                    desc = $"Filters a time intelligence result to only return historical dates, where {col.DaxObjectFullName} = {literal} (HF)";
                }
            }
            else if (methodB.Checked)
            {
                var col = (factColCombo.SelectedItem as ComboBoxItem)?.Column;
                if (col == null)
                {
                    Info("No transaction date column selected. Skipping 'Hide future dates'.");
                }
                else
                {
                    body = $@"(
    valueExpr: ANYREF EXPR
) =>
    // https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    VAR _LastVisibleDate  = MAX( {actualDateCol} )
    VAR _LastDateWithData =
        CALCULATE(
            MAX( {col.DaxObjectFullName} ),
            KEEPFILTERS( {actualDateCol} <= _LastVisibleDate )
        )
    RETURN
        CALCULATE(
            valueExpr,
            KEEPFILTERS( {actualDateCol} <= _LastDateWithData )
        )";
                    desc = $"Filters a time intelligence result to only return dates up to the last date with data in {col.DaxObjectFullName} (HF)";
                }
            }
            else // methodC
            {
                body = $@"(
    valueExpr: ANYREF EXPR
) =>
    // https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    VAR _LastVisibleDate = MIN( MAX( {actualDateCol} ), TODAY() )
    RETURN
        CALCULATE(
            valueExpr,
            KEEPFILTERS( {actualDateCol} <= _LastVisibleDate )
        )";
                desc = "Filters a time intelligence result to only return dates up to today (or the last selected date, whichever is earlier) (HF)";
            }

            if (body != null)
            {
                CreateUDF(hfFuncName, body, "HF", desc);
                hfFuncFullName = hfFuncName;
            }
        }
    }
}

// Wraps a raw calculation expression with the HF filter function for row i, if applicable
Func<int, string, string> WrapHF = (i, expr) =>
    (hideFutureSupportedIndices.Contains(i) && hideFutureChecks[i].Checked && hfFuncFullName != null)
        ? $"{hfFuncFullName}(\n        {expr}\n    )"
        : expr;

// Builds the "...to-date last period" pattern (PYTD/PQTD/PYWTD/PMTD/PWTD). For a plain Date
// column, nesting the shift function inside the ToDate function is correct and SQLBI-standard:
//   CALCULATE( baseMeasure, DATESYTD( SAMEPERIODLASTYEAR( dateColumn ) ) )
// For a Calendar option, that nesting does NOT work (DATEADD/SAMEPERIODLASTYEAR can't be nested
// inside DATESYTD/QTD/WTD/MTD when the argument is a Calendar object) - instead, the ToDate
// calculation must be computed FIRST (as an inner CALCULATE), then shifted via an OUTER CALCULATE
// using DATEADD as its filter argument (context transition), e.g.:
//   CALCULATE( CALCULATE( baseMeasure, DATESYTD( 'Gregorian' ) ), DATEADD( 'Gregorian', -1, YEAR ) )
Func<string, string, int, string, string> PreviousToDate = (todateFunc, dc, amount, unit) =>
{
    string shift = (!isCalendar && unit == "YEAR" && amount == -1)
        ? $"SAMEPERIODLASTYEAR( {dc} )"
        : $"DATEADD( {dc}, {amount}, {unit} )";

    return isCalendar
        ? $@"CALCULATE(
        CALCULATE( baseMeasure, {todateFunc}( {dc} ) ),
        {shift}
    )"
        : $@"CALCULATE(
        baseMeasure,
        {todateFunc}( {shift} )
    )";
};

// ===================== Create UDFs =====================

// 0 – Year-to-date
if (checks[0].Checked)
    CreateUDF(FuncName(suffixes[0].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(0, $@"CALCULATE(
        baseMeasure,
        DATESYTD( {dateColumn} )
    )")}", suffixes[0].Text, descriptions[0]);

// 1 – Year-to-date last year
if (checks[1].Checked)
    CreateUDF(FuncName(suffixes[1].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(1, PreviousToDate("DATESYTD", dateColumn, -1, "YEAR"))}", suffixes[1].Text, descriptions[1]);

// 2 – Quarter-to-date
if (checks[2].Checked)
    CreateUDF(FuncName(suffixes[2].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(2, $@"CALCULATE(
        baseMeasure,
        DATESQTD( {dateColumn} )
    )")}", suffixes[2].Text, descriptions[2]);

// 3 – Quarter-to-date last year
if (checks[3].Checked)
    CreateUDF(FuncName(suffixes[3].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(3, PreviousToDate("DATESQTD", dateColumn, -1, "YEAR"))}", suffixes[3].Text, descriptions[3]);

// 4 – Week-to-date
if (checks[4].Checked)
    CreateUDF(FuncName(suffixes[4].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(4, $@"CALCULATE(
        baseMeasure,
        DATESWTD( {dateColumn} )
    )")}", suffixes[4].Text, descriptions[4]);

// 5 – Week-to-date last year
if (checks[5].Checked)
    CreateUDF(FuncName(suffixes[5].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(5, PreviousToDate("DATESWTD", dateColumn, -1, "YEAR"))}", suffixes[5].Text, descriptions[5]);

// 6 – Week-to-date previous week
if (checks[6].Checked)
    CreateUDF(FuncName(suffixes[6].Text), $@"(
    //https://www.sqlbi.com/articles/understanding-dateadd-parameters-with-calendar-based-time-intelligence/
    baseMeasure: ANYREF
) =>
    {WrapHF(6, PreviousToDate("DATESWTD", dateColumn, -1, "WEEK"))}", suffixes[6].Text, descriptions[6]);

// 7 – Month-to-date
if (checks[7].Checked)
    CreateUDF(FuncName(suffixes[7].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(7, $@"CALCULATE(
        baseMeasure,
        DATESMTD( {dateColumn} )
    )")}", suffixes[7].Text, descriptions[7]);

// 8 – Month-to-date last year
if (checks[8].Checked)
    CreateUDF(FuncName(suffixes[8].Text), $@"(
    baseMeasure: ANYREF
) =>
    {WrapHF(8, PreviousToDate("DATESMTD", dateColumn, -1, "YEAR"))}", suffixes[8].Text, descriptions[8]);

// 9 – Previous year (full year total, accumulated) - uses the real column {actualDateCol} rather
// than {dateColumn}, so it works the same whether {dateColumn} is a plain column or a Calendar
// option (PARALLELPERIOD/DATEADD would not accept a Calendar option here).
if (checks[9].Checked)
    CreateUDF(FuncName(suffixes[9].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        FILTER(
            ALL( {actualDateCol} ),
            YEAR( {actualDateCol} ) = YEAR( MAX( {actualDateCol} ) ) - 1
        )
    )",
    suffixes[9].Text, descriptions[9]);

// 10 – Previous month (full month total, accumulated) - same {actualDateCol}-based approach as
// row 9, robust across year boundaries (e.g. January's previous month is December last year).
if (checks[10].Checked)
    CreateUDF(FuncName(suffixes[10].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        FILTER(
            ALL( {actualDateCol} ),
            YEAR( {actualDateCol} ) * 12 + MONTH( {actualDateCol} )
                = YEAR( MAX( {actualDateCol} ) ) * 12 + MONTH( MAX( {actualDateCol} ) ) - 1
        )
    )",
    suffixes[10].Text, descriptions[10]);

// 11 – Rolling N days
if (checks[11].Checked)
    CreateUDF(FuncName(suffixes[11].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _PreviousDates =
    DATESINPERIOD(
        {dateColumn},
        _ReferenceDate,
        -{nDays}, // can be replaced with _dynamicRollingPeriod
        DAY
    )
RETURN
    {WrapHF(11, $@"CALCULATE(
        baseMeasure,
        _PreviousDates
    )")}", suffixes[11].Text, descriptions[11]);

// 12 – Rolling N months
if (checks[12].Checked)
    CreateUDF(FuncName(suffixes[12].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _PreviousDates =
    DATESINPERIOD(
        {dateColumn},
        _ReferenceDate,
        -{nMonths}, // can be replaced with _dynamicRollingPeriod
        MONTH
    )
RETURN
    {WrapHF(12, $@"CALCULATE(
        baseMeasure,
        _PreviousDates
    )")}", suffixes[12].Text, descriptions[12]);

// 13 – Rolling N weeks
if (checks[13].Checked)
    CreateUDF(FuncName(suffixes[13].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _PreviousDates =
    DATESINPERIOD(
        {dateColumn},
        _ReferenceDate,
        -{nWeeks}, // can be replaced with _dynamicRollingPeriod
        WEEK
    )
RETURN
    {WrapHF(13, $@"CALCULATE(
        baseMeasure,
        _PreviousDates
    )")}", suffixes[13].Text, descriptions[13]);

// 14 – Rolling N months avg/month
if (checks[14].Checked)
    CreateUDF(FuncName(suffixes[14].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _LastCurrentDate = MAX( {actualDateCol})
VAR _Period =
    DATESINPERIOD(
        {dateColumn},
        _LastCurrentDate,
        -{nMonths}, // can be replaced with _dynamicRollingPeriod
        MONTH
    )
RETURN
    {WrapHF(14, $@"CALCULATE(
        AVERAGEX(
            VALUES( {yearMonthColumn} ),
            baseMeasure
        ),
        _Period
    )")}", suffixes[14].Text, descriptions[14]);

// 29 – Rolling N weeks avg/week
if (checks[RollingWeeksAvgIndex].Checked)
    CreateUDF(FuncName(suffixes[RollingWeeksAvgIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _LastCurrentDate = MAX( {actualDateCol})
VAR _Period =
    DATESINPERIOD(
        {dateColumn},
        _LastCurrentDate,
        -{nWeeks}, // can be replaced with _dynamicRollingPeriod
        WEEK
    )
RETURN
    {WrapHF(RollingWeeksAvgIndex, $@"CALCULATE(
        AVERAGEX(
            VALUES( {yearWeekColumn} ),
            baseMeasure
        ),
        _Period
    )")}", suffixes[RollingWeeksAvgIndex].Text, descriptions[RollingWeeksAvgIndex]);

// 15 – Rolling N days avg/day
if (checks[15].Checked)
    CreateUDF(FuncName(suffixes[15].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _LastCurrentDate = MAX( {actualDateCol} )
VAR _Period =
    DATESINPERIOD(
        {dateColumn},
        _LastCurrentDate,
        -{nDays}, // can be replaced with _dynamicRollingPeriod
        DAY
    )
RETURN
    {WrapHF(15, $@"CALCULATE(
        AVERAGEX(
            VALUES( {actualDateCol} ),
            baseMeasure
        ),
        _Period
    )")}", suffixes[15].Text, descriptions[15]);

// 30 – Rolling N days avg/transaction date
// Local/model-dependent observed-period variant based on:
// https://www.sqlbi.com/articles/the-hidden-complexity-of-a-daily-average-in-dax/
if (checks[RollingDaysObservedAvgIndex].Checked && observedTransactionDateColumn != null)
    CreateUDF(FuncName(suffixes[RollingDaysObservedAvgIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

//VAR _dynamicRollingPeriod = SELECTEDVALUE( 'Table'[Column] )
VAR _FirstObservedDate =
    CALCULATE(
        MIN( {observedTransactionDateColumn} ),
        REMOVEFILTERS()
    )
VAR _LastObservedDate =
    CALCULATE(
        MAX( {observedTransactionDateColumn} ),
        REMOVEFILTERS()
    )
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _RollingDates =
    DATESINPERIOD(
        {actualDateCol},
        _ReferenceDate,
        -{nDays}, // can be replaced with _dynamicRollingPeriod
        DAY
    )
VAR _ObservedDates =
    FILTER(
        _RollingDates,
        NOT ISBLANK( {actualDateCol} )
            && {actualDateCol} >= _FirstObservedDate
            && {actualDateCol} <= _LastObservedDate
    )
VAR _NumberOfDays = COUNTROWS( _ObservedDates )
VAR _Amount =
    CALCULATE(
        baseMeasure,
        KEEPFILTERS( _ObservedDates )
    )
RETURN
    IF(
        NOT ISBLANK( _FirstObservedDate )
            && NOT ISBLANK( _LastObservedDate )
            && _NumberOfDays > 0,
        DIVIDE(
            COALESCE( _Amount, 0 ),
            _NumberOfDays
        )
    )", suffixes[RollingDaysObservedAvgIndex].Text, descriptions[RollingDaysObservedAvgIndex]);

// 16 – Running total
if (checks[16].Checked)
    CreateUDF(FuncName(suffixes[16].Text), $@"(
    baseMeasure: ANYREF
) =>

VAR _currdate = MAX( {actualDateCol} )
RETURN
    {WrapHF(16, $@"CALCULATE(
        baseMeasure,
        FILTER(
            ALLSELECTED( {dateColumn} ),
            ISONORAFTER( {dateColumn}, _currdate, DESC )
        )
    )")}", suffixes[16].Text, descriptions[16]);

// 17 – Previous Period (standard Date column only - not available for a Calendar option, since it
// relies on DATESBETWEEN over the raw MIN/MAX date selection rather than a calendar time-unit function)
if (checks[17].Checked)
    CreateUDF(FuncName(suffixes[17].Text), $@"(
    baseMeasure: ANYREF
) =>

VAR _CurrentEndDate    = MAX( {actualDateCol} )
VAR _CurrentStartDate  = MIN( {actualDateCol} )
VAR _PeriodLength      = DATEDIFF( _CurrentStartDate, _CurrentEndDate, DAY )
VAR _PreviousEndDate   = _CurrentStartDate - 1
VAR _PreviousStartDate = _PreviousEndDate - _PeriodLength
RETURN
    CALCULATE(
        baseMeasure,
        DATESBETWEEN( {dateColumn}, _PreviousStartDate, _PreviousEndDate )
    )", suffixes[17].Text, descriptions[17]);

// 18 – Current year (full year total, accumulated) - same {actualDateCol}-based approach as
// row 9, always the whole current year regardless of the current filter granularity.
if (checks[18].Checked)
    CreateUDF(FuncName(suffixes[18].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        FILTER(
            ALL( {actualDateCol} ),
            YEAR( {actualDateCol} ) = YEAR( MAX( {actualDateCol} ) )
        )
    )",
    suffixes[18].Text, descriptions[18]);

// 19 – Rolling N months last year
if (checks[RollingMonthsLyIndex].Checked)
    CreateUDF(FuncName(suffixes[RollingMonthsLyIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _PreviousDates =
    DATESINPERIOD(
        {dateColumn},
        _ReferenceDate,
        -{nMonths}, // can be replaced with _dynamicRollingPeriod
        MONTH
    )
RETURN
    {WrapHF(RollingMonthsLyIndex, isCalendar
        // DATEADD (and SAMEPERIODLASTYEAR) can't take a VAR table expression like _PreviousDates
        // as its first argument when working with a Calendar object - it must reference the
        // Calendar itself directly. So the shift is applied as an OUTER CALCULATE (context
        // transition) around an INNER CALCULATE that restricts to _PreviousDates - same pattern
        // used by PreviousToDate() above for the ToDate calcs.
        ? $@"CALCULATE(
        CALCULATE( baseMeasure, _PreviousDates ),
        DATEADD( {dateColumn}, -1, YEAR )
    )"
        : $@"CALCULATE(
        baseMeasure,
        SAMEPERIODLASTYEAR( _PreviousDates )
    )")}", suffixes[RollingMonthsLyIndex].Text, descriptions[RollingMonthsLyIndex]);

// 20 – Rolling N months avg/month last year
if (checks[RollingMonthsAvgLyIndex].Checked)
    CreateUDF(FuncName(suffixes[RollingMonthsAvgLyIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _ReferenceDate = MAX( {actualDateCol} )
VAR _PreviousDates =
    DATESINPERIOD(
        {dateColumn},
        _ReferenceDate,
        -{nMonths}, // can be replaced with _dynamicRollingPeriod
        MONTH
    )
RETURN
    {WrapHF(RollingMonthsAvgLyIndex, isCalendar
        // Same reasoning as index 19 above: DATEADD needs the Calendar object itself, not the
        // _PreviousDates VAR, so the year-shift is applied as an outer CALCULATE (context
        // transition) wrapping an inner CALCULATE that computes the average over _PreviousDates.
        ? $@"CALCULATE(
        CALCULATE(
            AVERAGEX(
                VALUES( {yearMonthColumn} ),
                baseMeasure
            ),
            _PreviousDates
        ),
        DATEADD( {dateColumn}, -1, YEAR )
    )"
        : $@"CALCULATE(
        AVERAGEX(
            VALUES( {yearMonthColumn} ),
            baseMeasure
        ),
        SAMEPERIODLASTYEAR( _PreviousDates )
    )")}", suffixes[RollingMonthsAvgLyIndex].Text, descriptions[RollingMonthsAvgLyIndex]);

// 21 – Incoming balance
if (checks[IncomingBalanceIndex].Checked)
    CreateUDF(FuncName(suffixes[IncomingBalanceIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

VAR _earliestDate = MIN( {actualDateCol} )
RETURN
    {WrapHF(IncomingBalanceIndex, $@"CALCULATE(
        baseMeasure,
        FILTER(
            ALLSELECTED( {dateColumn} ),
            ISONORAFTER( {dateColumn}, _earliestDate, DESC )
        ),
        {actualDateCol} < _earliestDate
    )")}", suffixes[IncomingBalanceIndex].Text, descriptions[IncomingBalanceIndex]);

// 22 – LastDate
if (checks[LastDateIndex].Checked)
{
    if (!string.IsNullOrWhiteSpace(lastDateFactColumn) && !string.IsNullOrWhiteSpace(lastDateFactTableRef) && !string.IsNullOrWhiteSpace(dateTableRef))
        CreateUDF(FuncName(suffixes[LastDateIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxBalanceDate =
    CALCULATE(
        MAX( {lastDateFactColumn} ),  // Use MIN for Balance FirstDateWithData
        ALLEXCEPT( {lastDateFactTableRef}, {dateTableRef} )
    )
RETURN
    CALCULATE(
        baseMeasure,
        {actualDateCol} = _MaxBalanceDate
    )", suffixes[LastDateIndex].Text, descriptions[LastDateIndex]);
    else
        skipped.Add($"'{FuncName(suffixes[LastDateIndex].Text)}' was NOT created - no LastDate fact-table date column was selected.");
}

// 27 – LastDate per Category
if (checks[LastDatePerCategoryIndex].Checked)
{
    if (!string.IsNullOrWhiteSpace(lastDateFactColumn) && !string.IsNullOrWhiteSpace(dateTableRef))
        CreateUDF(FuncName(suffixes[LastDatePerCategoryIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX( {lastDateFactColumn} )
VAR _noDays = 1
VAR _ProdDates =
    CALCULATETABLE(
        VALUES( {lastDateFactColumn} ),
        {lastDateFactColumn} <= _MaxFactDate,
        REMOVEFILTERS( {dateTableRef} )
    )
VAR _LastProdDates =
    TOPN( _noDays, _ProdDates, {lastDateFactColumn}, DESC )
RETURN
    CALCULATE(
        baseMeasure,
        {actualDateCol} IN _LastProdDates
    )", suffixes[LastDatePerCategoryIndex].Text, descriptions[LastDatePerCategoryIndex]);
    else
        skipped.Add($"'{FuncName(suffixes[LastDatePerCategoryIndex].Text)}' was NOT created - no LastDate fact-table date column was selected.");
}

// 28 – Previous per Category
if (checks[PreviousPerCategoryIndex].Checked)
{
    if (!string.IsNullOrWhiteSpace(lastDateFactColumn) && !string.IsNullOrWhiteSpace(dateTableRef))
        CreateUDF(FuncName(suffixes[PreviousPerCategoryIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX( {lastDateFactColumn} )
VAR _noDays = 1
VAR _noDaysPrevious = _noDays * 2
VAR _ProdDates =
    CALCULATETABLE(
        VALUES( {lastDateFactColumn} ),
        {lastDateFactColumn} <= _MaxFactDate,
        REMOVEFILTERS( {dateTableRef} )
    )
VAR _PreviousProdDates =
    TOPN( _noDaysPrevious, _ProdDates, {lastDateFactColumn}, DESC )
VAR _LastProdDates =
    TOPN( _noDays, _ProdDates, {lastDateFactColumn}, DESC )
RETURN
    CALCULATE(
        baseMeasure,
        {actualDateCol} IN _PreviousProdDates,
        NOT {actualDateCol} IN _LastProdDates
    )", suffixes[PreviousPerCategoryIndex].Text, descriptions[PreviousPerCategoryIndex]);
    else
        skipped.Add($"'{FuncName(suffixes[PreviousPerCategoryIndex].Text)}' was NOT created - no LastDate fact-table date column was selected.");
}

// 23 – Open - End of period
if (checks[OpenEndOfPeriodIndex].Checked)
{
    if (!string.IsNullOrWhiteSpace(fromDateCol) && !string.IsNullOrWhiteSpace(toDateCol) && !string.IsNullOrWhiteSpace(dateTableRef))
        CreateUDF(FuncName(suffixes[OpenEndOfPeriodIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.daxpatterns.com/semi-additive-calculations/
// https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = MAX( {actualDateCol} )
RETURN
    CALCULATE(
        baseMeasure,
        REMOVEFILTERS( {dateTableRef} ),
        {fromDateCol} <= _EndDateVisual,
        {toDateCol} > _EndDateVisual || ISBLANK( {toDateCol} )
    )", suffixes[OpenEndOfPeriodIndex].Text, descriptions[OpenEndOfPeriodIndex]);
    else
        skipped.Add($"'{FuncName(suffixes[OpenEndOfPeriodIndex].Text)}' was NOT created - no from/to date columns were selected for the open-period calculation.");
}

// 24 – Total during period
if (checks[TotalDuringPeriodIndex].Checked)
{
    if (!string.IsNullOrWhiteSpace(fromDateCol) && !string.IsNullOrWhiteSpace(toDateCol) && !string.IsNullOrWhiteSpace(dateTableRef))
        CreateUDF(FuncName(suffixes[TotalDuringPeriodIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.daxpatterns.com/semi-additive-calculations/
// https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = MAX( {actualDateCol} )
VAR _StartDateVisual = MIN( {actualDateCol} )
RETURN
    CALCULATE(
        baseMeasure,
        REMOVEFILTERS( {dateTableRef} ),
        {fromDateCol} <= _EndDateVisual,
        {toDateCol} > _StartDateVisual || ISBLANK( {toDateCol} )
    )", suffixes[TotalDuringPeriodIndex].Text, descriptions[TotalDuringPeriodIndex]);
    else
        skipped.Add($"'{FuncName(suffixes[TotalDuringPeriodIndex].Text)}' was NOT created - no from/to date columns were selected for the open-period calculation.");
}

// 25 – Simple filter - Historical dates
if (checks[HistoricalDatesIndex].Checked)
    CreateUDF(FuncName(suffixes[HistoricalDatesIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.sqlbi.com/articles/showing-actuals-and-forecasts-in-the-same-chart-with-power-bi/
VAR _ReferenceDate = TODAY()
RETURN
    IF( MAX( {actualDateCol} ) <= _ReferenceDate, baseMeasure )", suffixes[HistoricalDatesIndex].Text, descriptions[HistoricalDatesIndex]);

// 26 – Simple filter - Remaining year
if (checks[RemainingYearIndex].Checked)
    CreateUDF(FuncName(suffixes[RemainingYearIndex].Text), $@"(
    baseMeasure: ANYREF
) =>

// https://www.sqlbi.com/articles/showing-actuals-and-forecasts-in-the-same-chart-with-power-bi/
VAR _ReferenceDate = TODAY()
RETURN
    IF( MIN( {actualDateCol} ) > _ReferenceDate, baseMeasure )", suffixes[RemainingYearIndex].Text, descriptions[RemainingYearIndex]);

// ===================== Fiscal Year (optional add-on, YTD only) =====================
// DATESYTD is the only classic time-intelligence function that accepts a fiscal year-end
// date, so the Fiscal Year option only produces YTD / YTD-last-year (PYTD) variants.
int _fiscalUdfCount = 0;
if (fiscalYearCheckbox.Checked)
{
    string _fiscalYearInput = Interaction.InputBox(
        "FISCAL YEAR INPUT:" + Environment.NewLine + "Last day of the fiscal year",
        "Fiscal year input, last date MM/DD", "06/30");

    if (string.IsNullOrEmpty(_fiscalYearInput))
    {
        Info("No fiscal year input provided. Skipping Fiscal Year UDFs.");
    }
    else
    {
        CreateUDF(FuncName("YTD_FY"), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESYTD( {dateColumn}, ""{_fiscalYearInput}"" )
    )", "YTD_FY", $"Accumulated fiscal year to date, fiscal year ending {_fiscalYearInput} (YTD FY)");

        CreateUDF(FuncName("PYTD_FY"), $@"(
    baseMeasure: ANYREF
) =>
    {(isCalendar
        ? $@"CALCULATE(
        CALCULATE( baseMeasure, DATESYTD( {dateColumn}, ""{_fiscalYearInput}"" ) ),
        DATEADD( {dateColumn}, -1, YEAR )
    )"
        : $@"CALCULATE(
        baseMeasure,
        DATESYTD( SAMEPERIODLASTYEAR( {dateColumn} ), ""{_fiscalYearInput}"" )
    )")}", "PYTD_FY", $"Previous fiscal year to date, fiscal year ending {_fiscalYearInput} (PYTD FY)");

        _fiscalUdfCount = 2;
        Info($"Created Fiscal Year variants (TI_YTD_FY / TI_PYTD_FY) using fiscal year-end {_fiscalYearInput}.");
    }
}

// ===================== Comparison (Dev / idx) add-on =====================
// For the 6 supported rows (YTD, QTD, WTD, MTD, RNm, RNm avg), if the "Cmp" box was checked,
// creates two extra UDFs comparing the calc's own value ("actual") against its natural
// prior-period counterpart ("target": PYTD / PQTD / PYWTD / PMTD / RNm LY / RNm avg LY
// respectively), using the standard model-independent Comparison.Deviation / Comparison.Pct_change
// UDFs (created here if they don't already exist) - same functions/pattern as
// "5 - Compare first selected measure with second selected measure.cs", limited to just these two.
int _comparisonUdfCount = 0;
if (comparisonSupportedIndices.Any(i => checks[i].Checked && comparisonChecks[i].Checked))
{
    const string devFuncName = "Comparison.Deviation";
    const string idxFuncName = "Comparison.Pct_change";

    if (!Model.Functions.Any(f => f.Name == devFuncName))
    {
        var devFn = Model.AddFunction(devFuncName);
        devFn.Expression = @"(
    // What you are measuring
    actual,

    // What you want to compare to
    target
) =>
    VAR _Target = target
    VAR _Delta = actual - _Target

    RETURN
        _Delta";
        devFn.Description = "Relative difference between measure1 and measure2.";
    }

    if (!Model.Functions.Any(f => f.Name == idxFuncName))
    {
        var idxFn = Model.AddFunction(idxFuncName);
        idxFn.Expression = @"(
    // What you are measuring
    actual,

    // What you want to compare to
    target
) =>
    VAR _Target = target
    VAR _Delta = actual - _Target
    VAR _Result = DIVIDE ( _Delta, ABS ( _Target ) )
    RETURN
        _Result";
        idxFn.Description = "Relative percentage difference between measure1 and measure2.";
    }

    // Builds "<suffix> Dev" and "<suffix> idx" UDFs comparing actualExpr vs. targetExpr (each a
    // full DAX expression, possibly its own nested VAR/RETURN block, that resolves to a scalar
    // using baseMeasure).
    Action<int, string, string> AddComparisonPair = (sourceIdx, actualExpr, targetExpr) =>
    {
        if (!(checks[sourceIdx].Checked && comparisonChecks[sourceIdx].Checked)) return;

        string baseSuffix = suffixes[sourceIdx].Text;
        string devSuffix  = baseSuffix + " Dev";
        string idxSuffix  = baseSuffix + " idx";

        CreateUDF(FuncName(devSuffix), $@"(
    baseMeasure: ANYREF
) =>
    VAR _Actual = {actualExpr}
    VAR _Target = {targetExpr}
    RETURN
        {devFuncName}( _Actual, _Target )",
        devSuffix, $"Absolute deviation of {baseSuffix} vs. its prior-period counterpart, using {devFuncName} ({devSuffix}).");

        CreateUDF(FuncName(idxSuffix), $@"(
    baseMeasure: ANYREF
) =>
    VAR _Actual = {actualExpr}
    VAR _Target = {targetExpr}
    RETURN
        {idxFuncName}( _Actual, _Target )",
        idxSuffix, $"Percentage change of {baseSuffix} vs. its prior-period counterpart, using {idxFuncName} ({idxSuffix}).");

        _comparisonUdfCount += 2;
    };

    // YTD vs PYTD
    AddComparisonPair(0,
        WrapHF(0, $@"CALCULATE(
        baseMeasure,
        DATESYTD( {dateColumn} )
    )"),
        WrapHF(1, PreviousToDate("DATESYTD", dateColumn, -1, "YEAR")));

    // QTD vs PQTD
    AddComparisonPair(2,
        WrapHF(2, $@"CALCULATE(
        baseMeasure,
        DATESQTD( {dateColumn} )
    )"),
        WrapHF(3, PreviousToDate("DATESQTD", dateColumn, -1, "YEAR")));

    // WTD vs PYWTD (previous year, not the "previous week" PWTD variant)
    AddComparisonPair(4,
        WrapHF(4, $@"CALCULATE(
        baseMeasure,
        DATESWTD( {dateColumn} )
    )"),
        WrapHF(5, PreviousToDate("DATESWTD", dateColumn, -1, "YEAR")));

    // MTD vs PMTD
    AddComparisonPair(7,
        WrapHF(7, $@"CALCULATE(
        baseMeasure,
        DATESMTD( {dateColumn} )
    )"),
        WrapHF(8, PreviousToDate("DATESMTD", dateColumn, -1, "YEAR")));

    // RNm vs RNm LY
    AddComparisonPair(12,
        WrapHF(12, $@"VAR _ReferenceDate = MAX( {actualDateCol} )
    VAR _PreviousDates =
        DATESINPERIOD(
            {dateColumn},
            _ReferenceDate,
            -{nMonths},
            MONTH
        )
    RETURN
        CALCULATE(
            baseMeasure,
            _PreviousDates
        )"),
        WrapHF(RollingMonthsLyIndex, $@"VAR _ReferenceDate = MAX( {actualDateCol} )
    VAR _PreviousDates =
        DATESINPERIOD(
            {dateColumn},
            _ReferenceDate,
            -{nMonths},
            MONTH
        )
    RETURN
        CALCULATE(
            baseMeasure,
            {(isCalendar ? "DATEADD( _PreviousDates, -1, YEAR )" : "SAMEPERIODLASTYEAR( _PreviousDates )")}
        )"));

    // RNm avg vs RNm avg LY
    AddComparisonPair(14,
        WrapHF(14, $@"VAR _LastCurrentDate = MAX( {actualDateCol} )
    VAR _Period =
        DATESINPERIOD(
            {dateColumn},
            _LastCurrentDate,
            -{nMonths},
            MONTH
        )
    RETURN
        CALCULATE(
            AVERAGEX(
                VALUES( {yearMonthColumn} ),
                baseMeasure
            ),
            _Period
        )"),
        WrapHF(RollingMonthsAvgLyIndex, $@"VAR _ReferenceDate = MAX( {actualDateCol} )
    VAR _PreviousDates =
        DATESINPERIOD(
            {dateColumn},
            _ReferenceDate,
            -{nMonths},
            MONTH
        )
    VAR _ShiftedPeriod =
        {(isCalendar ? "DATEADD( _PreviousDates, -1, YEAR )" : "SAMEPERIODLASTYEAR( _PreviousDates )")}
    RETURN
        CALCULATE(
            AVERAGEX(
                VALUES( {yearMonthColumn} ),
                baseMeasure
            ),
            _ShiftedPeriod
        )"));
}

string doneSummary = "Done! " + (checks.Count(c => c.Checked) + _fiscalUdfCount + _comparisonUdfCount) + " UDF(s) requested under prefix '" + namePrefix + "'.";
if (skipped.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, skipped.Select(s => "- " + s));
Info(doneSummary);

// ===================== Helper class =====================
class ComboBoxItem
{
    public Column   Column   { get; private set; }
    public Calendar Calendar { get; private set; }
    public Table    Table    { get; private set; }
    public ComboBoxItem(Column   column)   { this.Column   = column;   }
    public ComboBoxItem(Calendar calendar) { this.Calendar = calendar; }
    public ComboBoxItem(Table    table)    { this.Table    = table;    }
    public override string ToString() =>
        Table != null ? Table.Name :
        Calendar == null
            ? "Column: "   + Column.DaxObjectFullName
            : "Calendar: " + Calendar.DaxObjectFullName;
}