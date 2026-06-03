// ADD TIME INTELLIGENCE AS DAX UDFs
// ===================================
// Inspired by: https://github.com/TabularEditor/SessionArchive/blob/main/Webinars/2026-02-04-scripting-automation/add-time-intelligence.csx
//
// PURPOSE
// -------
// Creates reusable DAX User-Defined Functions (UDFs) for time intelligence
// directly in the semantic model. Each UDF wraps a standard DAX calculation
// pattern and accepts a single parameter (baseMeasure: ANYREF), so any measure
// can later be passed in — e.g. [Local.TimeIntelligence.TI_YTD]([Sales Amount]).
//
// HOW IT WORKS
// ------------
// 1. Run the script with no measure selection required (unlike the original
//    Tabular Editor sample script that generates measures directly).
// 2. Choose which Date column or Calendar to bind the UDFs to (see below).
// 3. Pick the time intelligence calculations you want (YTD, MTD, rolling, etc.).
// 4. Click OK — the script creates DAX UDF objects in the model, one per
//    selected calculation.
//
// STANDARD (DATE COLUMN) vs CALENDAR TIME INTELLIGENCE
// -----------------------------------------------------
// Standard — select a plain DateTime column on a table marked as a Date Table
//   (e.g. Calendar[Date]).  All standard DAX time-intelligence functions
//   (DATESYTD, DATESMTD, DATESQTD, PARALLELPERIOD, SAMEPERIODLASTYEAR, etc.)
//   work out-of-the-box because they rely on the continuous date table contract.
//   Week-based calculations are NOT available in this mode.
//
// Calendar — select a named Calendar object defined in the model (TE3 feature).
//   A Calendar encodes fiscal/custom time units (weeks, fiscal months, fiscal
//   quarters, broadcasting periods, etc.).  The script inspects the calendar's
//   time units and shows only the calculations that are valid for that calendar:
//   • Weekly calendars  → WTD / WTD LY / WTD PW / Rolling N weeks only.
//   • Monthly calendars → YTD / QTD / MTD / rolling months/days, etc.
//   Calendar-aware UDFs use DATESWTD, DATESQTD, DATESMTD, DATESYTD and DATEADD
//   with the calendar's DAX name instead of a raw column reference.
//
// ANNOTATIONS SET ON EACH UDF
// ---------------------------
// Every generated function receives three annotations that a companion macro
// can read later to auto-create measures for selected base measures:
//
//   FunctionCategory  = "Time Intelligence"
//       Identifies the function as a TI UDF (used for filtering / discovery).
//
//   MeasureSuffix     = e.g. "YTD", "MTD", "PY", "R12m" …
//       The suffix to append to the base measure name when a downstream macro
//       generates new measures:  [Sales Amount] → [Sales Amount YTD].
//
//   FolderSuffix      = e.g. "Time Intelligence"
//       The display-folder suffix appended to the base measure's own folder,
//       so generated measures land in  "Revenue\Time Intelligence"  etc.
//
// COMPANION MACRO (NEXT STEP)
// ---------------------------
// A separate script can enumerate all model functions where
// GetAnnotation("FunctionCategory") == "Time Intelligence", then loop over
// every selected measure and create a new measure per UDF using the
// MeasureSuffix and FolderSuffix annotations — giving you a fully automated
// time-intelligence measure factory without touching the UDFs again.

#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;

Application.UseWaitCursor = false;
WaitFormVisible = false;

var aggForm = new Form
{
    Text = "Add Time Intelligence UDFs",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = false,
    Width = 780,
    Height = 850,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterParent
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

var checks       = new List<CheckBox>();
var suffixes     = new List<TextBox>();
var descriptions = new List<string>();

Action<string, string, string> AddCalc = (label, defaultSuffix, description) => {
    var panel    = new Panel { AutoSize = true };
    var checkBox = new CheckBox { Text = label, AutoSize = true, Location = new Point(0, 2), Checked = false };
    panel.Controls.Add(checkBox);
    checks.Add(checkBox);
    var textBox = new TextBox { Location = new Point(400, 0), Width = 200, Text = defaultSuffix };
    suffixes.Add(textBox);
    panel.Controls.Add(textBox);
    flowPanel.Controls.Add(panel);
    descriptions.Add(description);
};

// ── 1. Date column / calendar selector ───────────────────────────────────
AddLabel("Choose Date column or calendar:");
var dateColumnSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = true, Width = 350 };

var dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime && c.IsKey).ToList();
var calendars   = Model.AllCalendars.ToList();
if (dateColumns.Count + calendars.Count == 0) {
    Info("Model contains no Date columns or Calendars.");
    return;
}
dateColumns.ForEach(c => dateColumnSelector.Items.Add(new ComboBoxItem(c)));
calendars.ForEach(c   => dateColumnSelector.Items.Add(new ComboBoxItem(c)));
dateColumnSelector.SelectedIndex = 0;
flowPanel.Controls.Add(dateColumnSelector);

// ── 2. Actual date column selector (used inside VALUES / AVERAGEX) ────────
AddLabel("Choose actual Date column (used inside VALUES for rolling avg, must be a column):");
var actualDateColLabel    = flowPanel.Controls[flowPanel.Controls.Count - 1] as Label;
var actualDateColSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = true, Width = 350 };
dateColumns.ForEach(c => actualDateColSelector.Items.Add(new ComboBoxItem(c)));
if (actualDateColSelector.Items.Count > 0) actualDateColSelector.SelectedIndex = 0;
flowPanel.Controls.Add(actualDateColSelector);

// ── 3. YearMonthNo column selector ───────────────────────────────────────
AddLabel("Choose YearMonthNo column (integer sort key, used for rolling averages):");
var yearMonthLabel    = flowPanel.Controls[flowPanel.Controls.Count - 1] as Label;
var yearMonthSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = true, Width = 350 };
var intColumns = Model.AllColumns
    .Where(c => c.DataType == DataType.Int64 && c.Name.IndexOf("month", StringComparison.OrdinalIgnoreCase) >= 0)
    .ToList();
if (intColumns.Count == 0)
    intColumns = Model.AllColumns.Where(c => c.DataType == DataType.Int64).ToList();
intColumns.ForEach(c => yearMonthSelector.Items.Add(new ComboBoxItem(c)));
if (yearMonthSelector.Items.Count > 0) yearMonthSelector.SelectedIndex = 0;
flowPanel.Controls.Add(yearMonthSelector);

// ── 4. Calculation checkboxes ─────────────────────────────────────────────
AddLabel("Choose Time Intelligence calculations and function name suffixes:");

AddCalc("Agg. year-to-date",                  "YTD",
        "Accumulated year to date (YTD)");                                       // 0
AddCalc("Agg. year-to-date last year",         "LYTD",
        "Accumulated year to date last year (LYTD)");                            // 1
AddCalc("Agg. quarter-to-date",                "QTD",
        "Accumulated quarter to date (QTD)");                                    // 2
AddCalc("Agg. quarter-to-date last year",      "LQTD",
        "Accumulated quarter to date last year (LQTD)");                         // 3
AddCalc("Agg. week-to-date",                   "WTD",
        "Accumulated week to date (WTD)");                                       // 4
AddCalc("Agg. week-to-date last year",         "WTD LY",
        "Accumulated week to date last year (WTD LY)");                          // 5
AddCalc("Agg. week-to-date previous week",     "WTD PW",
        "Accumulated week to date previous week (WTD PW)");                      // 6
AddCalc("Agg. month-to-date",                  "MTD",
        "Accumulated month to date (MTD)");                                      // 7
AddCalc("Agg. month-to-date last year",        "LMTD",
        "Accumulated month to date last year (LMTD)");                           // 8
AddCalc("1 year prior",                        "PY",
        "Same period 1 year prior (PY)");                                        // 9
AddCalc("2 years prior",                       "2YP",
        "Same period 2 years prior (2YP)");                                      // 10
AddCalc("Rolling N days",                      "RNd",
        "Rolling period of N days, from reference date backwards");              // 11
AddCalc("Rolling N months",                    "RNm",
        "Rolling period of N months, from reference date backwards");            // 12
AddCalc("Rolling N weeks",                     "RNw",
        "Rolling period of N weeks, from reference date backwards");             // 13
AddCalc("Rolling N months avg/month",          "RNm avg",
        "Rolling N months average per month, from reference date backwards");    // 14
AddCalc("Rolling N days avg/day",              "RNd avg",
        "Rolling N days average per day, from reference date backwards");        // 15
AddCalc("Running total",                       "RT",
        "Running total, from the first date available until selected date");     // 16

// ── 5. Rolling N inputs ───────────────────────────────────────────────────
AddLabel("Rolling N days  (default 365 — used for Rolling N days and Rolling N days avg/day):");
var rollingDaysInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 3650, Value = 365 };
flowPanel.Controls.Add(rollingDaysInput);

AddLabel("Rolling N months  (default 12 — used for Rolling N months and Rolling N months avg/month):");
var rollingMonthsLabel = flowPanel.Controls[flowPanel.Controls.Count - 1] as Label;
var rollingMonthsInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 120, Value = 12 };
flowPanel.Controls.Add(rollingMonthsInput);

AddLabel("Rolling N weeks  (used for Rolling N weeks):");
var rollingWeeksLabel = flowPanel.Controls[flowPanel.Controls.Count - 1] as Label;
var rollingWeeksInput = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 520, Value = 4 };
flowPanel.Controls.Add(rollingWeeksInput);

// ── 6. Name / folder options ──────────────────────────────────────────────
AddLabel("(Optional) Function name prefix  [default: Local.TimeIntelligence]:");
var prefixTextbox = new TextBox { Width = 350, Text = "Local.TimeIntelligence" };
flowPanel.Controls.Add(prefixTextbox);

AddLabel("(Optional) FolderSuffix annotation value  [default: Time Intelligence]:");
var folderSuffixTextbox = new TextBox { Width = 350, Text = "Time Intelligence" };
flowPanel.Controls.Add(folderSuffixTextbox);

// ── 7. Buttons ────────────────────────────────────────────────────────────
var buttonPanel  = new Panel { Margin = new Padding(0,20,0,0), AutoSize = true };
var okButton     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Location = new Point(195, 0) };
var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel, Location = new Point(280, 0) };
buttonPanel.Controls.Add(okButton);
buttonPanel.Controls.Add(cancelButton);
flowPanel.Controls.Add(buttonPanel);

aggForm.AcceptButton = okButton;
aggForm.CancelButton = cancelButton;

// ── CheckCalc ─────────────────────────────────────────────────────────────
Action CheckCalc = () => {
    if (checks.Count < 17) return;

    var selItem    = dateColumnSelector.SelectedItem as ComboBoxItem;
    var cal        = selItem.Calendar;
    bool isWeekly  = cal != null && cal.GetTimeUnits().Any(t =>
                         t.TimeUnit is TimeUnit.Week or TimeUnit.WeekOfMonth
                                    or TimeUnit.WeekOfYear or TimeUnit.WeekOfQuarter);
    bool hasMonth   = cal == null || cal.GetTimeUnits().Any(t =>
                         t.TimeUnit is TimeUnit.Month or TimeUnit.MonthOfQuarter or TimeUnit.MonthOfYear);
    bool hasQuarter = cal == null || cal.GetTimeUnits().Any(t =>
                         t.TimeUnit is TimeUnit.Quarter or TimeUnit.QuarterOfYear);

    // Show actual date column selector only when a calendar (not a column) is selected
    bool isCalendar = cal != null;
    actualDateColLabel.Visible    = isCalendar;
    actualDateColSelector.Visible = isCalendar;

    // Reset all
    for (int i = 0; i < checks.Count; i++) checks[i].Checked = false;

    if (isWeekly) {
        checks[4].Checked  = true;   // WTD
        checks[5].Checked  = true;   // WTD LY
        checks[6].Checked  = true;   // WTD PW
        checks[13].Checked = true;   // RNw
    } else {
        checks[0].Checked  = true;   // YTD
        checks[1].Checked  = true;   // LYTD
        if (hasQuarter) { checks[2].Checked = true; checks[3].Checked = true; }
        if (hasMonth)   { checks[7].Checked = true; checks[8].Checked = true; }
        checks[11].Checked = true;   // RNd
        checks[12].Checked = true;   // RNm
    }

    // Week-only: visible + enabled only for weekly calendars
    checks[4].Parent.Visible  = isWeekly;
    checks[5].Parent.Visible  = isWeekly;
    checks[6].Parent.Visible  = isWeekly;
    checks[13].Parent.Visible = isWeekly;
    rollingWeeksLabel.Visible  = isWeekly;
    rollingWeeksInput.Visible  = isWeekly;

    // Quarter-dependent
    checks[2].Parent.Visible = hasQuarter;
    checks[3].Parent.Visible = hasQuarter;
    if (!hasQuarter) { checks[2].Checked = false; checks[3].Checked = false; }

    // Month-dependent
    checks[7].Parent.Visible = hasMonth;
    checks[8].Parent.Visible = hasMonth;
    if (!hasMonth) { checks[7].Checked = false; checks[8].Checked = false; }

    // Month-based rolling: hidden for weekly calendar
    bool showMonthRolling = !isWeekly;
    checks[12].Parent.Visible = showMonthRolling;
    checks[14].Parent.Visible = showMonthRolling;
    checks[15].Parent.Visible = showMonthRolling;
    if (!showMonthRolling) { checks[12].Checked = false; checks[14].Checked = false; checks[15].Checked = false; }

    yearMonthLabel.Visible     = showMonthRolling;
    yearMonthSelector.Visible  = showMonthRolling;
    rollingMonthsLabel.Visible = showMonthRolling;
    rollingMonthsInput.Visible = showMonthRolling;
};

dateColumnSelector.SelectionChangeCommitted += (s, e) => CheckCalc();
CheckCalc();

if (aggForm.ShowDialog() == DialogResult.Cancel) return;

// ===================== Resolve selections =====================
var selItem      = dateColumnSelector.SelectedItem as ComboBoxItem;
var dateColumn   = selItem.Column == null
                     ? selItem.Calendar.DaxObjectFullName   // e.g. 'Gregorian' or 'Weekly'
                     : selItem.Column.DaxObjectFullName;    // e.g. Calendar[Date]

// actualDateCol is always a real column — used inside VALUES() and AVERAGEX()
// When a plain date column was selected (not a calendar), it equals dateColumn
var actualDateCol = (selItem.Column != null)
                      ? selItem.Column.DaxObjectFullName
                      : (actualDateColSelector.SelectedItem as ComboBoxItem).Column.DaxObjectFullName;

var yearMonthColumn = yearMonthSelector.SelectedItem != null
                        ? (yearMonthSelector.SelectedItem as ComboBoxItem).Column.DaxObjectFullName
                        : "Calendar[Year Month No]";

var namePrefix   = string.IsNullOrWhiteSpace(prefixTextbox.Text)
                     ? "Local.TimeIntelligence"
                     : prefixTextbox.Text.Trim();

var folderSuffix = folderSuffixTextbox.Text.Trim();

var nDays   = (int)rollingDaysInput.Value;
var nMonths = (int)rollingMonthsInput.Value;
var nWeeks  = (int)rollingWeeksInput.Value;

suffixes[11].Text = suffixes[11].Text.Replace("N", nDays.ToString());
suffixes[12].Text = suffixes[12].Text.Replace("N", nMonths.ToString());
suffixes[13].Text = suffixes[13].Text.Replace("N", nWeeks.ToString());
suffixes[14].Text = suffixes[14].Text.Replace("N", nMonths.ToString());
suffixes[15].Text = suffixes[15].Text.Replace("N", nDays.ToString());
descriptions[11]  = descriptions[11].Replace("N", nDays.ToString());
descriptions[12]  = descriptions[12].Replace("N", nMonths.ToString());
descriptions[13]  = descriptions[13].Replace("N", nWeeks.ToString());
descriptions[14]  = descriptions[14].Replace("N", nMonths.ToString());
descriptions[15]  = descriptions[15].Replace("N", nDays.ToString());

Func<string, string> FuncName = suffix =>
    namePrefix + ".TI_" + suffix.Replace(" ", "_");

Action<string, string, string, string> CreateUDF = (funcName, daxBody, measureSuffix, description) =>
{
    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null) existing.Delete();
    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
    fn.SetAnnotation("FunctionCategory", "Time Intelligence");
    fn.SetAnnotation("MeasureSuffix",    measureSuffix);
    fn.SetAnnotation("FolderSuffix",     folderSuffix);
};

// ===================== Create UDFs =====================

// 0 – Year-to-date
if (checks[0].Checked)
    CreateUDF(FuncName(suffixes[0].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESYTD( {dateColumn} )
    )", suffixes[0].Text, descriptions[0]);

// 1 – Year-to-date last year
if (checks[1].Checked)
    CreateUDF(FuncName(suffixes[1].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESYTD( {dateColumn} ),
        SAMEPERIODLASTYEAR( {dateColumn} )
    )", suffixes[1].Text, descriptions[1]);

// 2 – Quarter-to-date
if (checks[2].Checked)
    CreateUDF(FuncName(suffixes[2].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESQTD( {dateColumn} )
    )", suffixes[2].Text, descriptions[2]);

// 3 – Quarter-to-date last year
if (checks[3].Checked)
    CreateUDF(FuncName(suffixes[3].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESQTD( {dateColumn} ),
        SAMEPERIODLASTYEAR( {dateColumn} )
    )", suffixes[3].Text, descriptions[3]);

// 4 – Week-to-date
if (checks[4].Checked)
    CreateUDF(FuncName(suffixes[4].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESWTD( {dateColumn} )
    )", suffixes[4].Text, descriptions[4]);

// 5 – Week-to-date last year
if (checks[5].Checked)
    CreateUDF(FuncName(suffixes[5].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESWTD( {dateColumn} ),
        SAMEPERIODLASTYEAR( {dateColumn} )
    )", suffixes[5].Text, descriptions[5]);

// 6 – Week-to-date previous week
if (checks[6].Checked)
    CreateUDF(FuncName(suffixes[6].Text), $@"(
    //https://www.sqlbi.com/articles/understanding-dateadd-parameters-with-calendar-based-time-intelligence/
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESWTD( {dateColumn} ),
        DATEADD( {dateColumn}, -1, WEEK )
    )", suffixes[6].Text, descriptions[6]);

// 7 – Month-to-date
if (checks[7].Checked)
    CreateUDF(FuncName(suffixes[7].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESMTD( {dateColumn} )
    )", suffixes[7].Text, descriptions[7]);

// 8 – Month-to-date last year
if (checks[8].Checked)
    CreateUDF(FuncName(suffixes[8].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE(
        baseMeasure,
        DATESMTD( {dateColumn} ),
        SAMEPERIODLASTYEAR( {dateColumn} )
    )", suffixes[8].Text, descriptions[8]);

// 9 – 1 year prior
if (checks[9].Checked)
    CreateUDF(FuncName(suffixes[9].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE( baseMeasure, PARALLELPERIOD( {dateColumn}, -1, YEAR ) )",
    suffixes[9].Text, descriptions[9]);

// 10 – 2 years prior
if (checks[10].Checked)
    CreateUDF(FuncName(suffixes[10].Text), $@"(
    baseMeasure: ANYREF
) =>
    CALCULATE( baseMeasure, PARALLELPERIOD( {dateColumn}, -2, YEAR ) )",
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
    CALCULATE(
        baseMeasure,
        _PreviousDates
    )", suffixes[11].Text, descriptions[11]);

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
    CALCULATE(
        baseMeasure,
        _PreviousDates
    )", suffixes[12].Text, descriptions[12]);

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
    CALCULATE(
        baseMeasure,
        _PreviousDates
    )", suffixes[13].Text, descriptions[13]);

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
    CALCULATE(
        AVERAGEX(
            VALUES( {yearMonthColumn} ),
            baseMeasure
        ),
        _Period
    )", suffixes[14].Text, descriptions[14]);

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
    CALCULATE(
        AVERAGEX(
            VALUES( {actualDateCol} ),
            baseMeasure
        ),
        _Period
    )", suffixes[15].Text, descriptions[15]);

// 16 – Running total
if (checks[16].Checked)
    CreateUDF(FuncName(suffixes[16].Text), $@"(
    baseMeasure: ANYREF
) =>

VAR _currdate = MAX( {actualDateCol} )
RETURN
    CALCULATE(
        baseMeasure,
        FILTER(
            ALLSELECTED( {actualDateCol} ),
            ISONORAFTER( {actualDateCol}, _currdate, DESC )
        )
    )", suffixes[16].Text, descriptions[16]);

Info("Done! " + checks.Count(c => c.Checked) + " UDF(s) created under prefix '" + namePrefix + "'.");

// ===================== Helper class =====================
class ComboBoxItem
{
    public Column   Column   { get; private set; }
    public Calendar Calendar { get; private set; }
    public ComboBoxItem(Column   column)   { this.Column   = column;   }
    public ComboBoxItem(Calendar calendar) { this.Calendar = calendar; }
    public override string ToString() =>
        Calendar == null
            ? "Column: "   + Column.DaxObjectFullName
            : "Calendar: " + Calendar.DaxObjectFullName;
}