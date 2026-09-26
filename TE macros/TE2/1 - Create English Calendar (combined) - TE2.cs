// The purpose of this macro is to create an English calendar table, mark it as a date table,
// and configure calendar options (Gregorian + ISO/Weekly) - all in a single script execution
// (no reload required between steps, since everything runs against the same in-memory model).
//
// This replaces the previous two-script workflow:
//   1 - Create English Calendar.cs
//   2 - Mark as Date-Table and create relationship.cs
//
// Changes vs. the original two scripts:
//   1. "Year Slicer" column renamed to "Year slicer" (matches the actual M-generated column name
//      produced by the RenameColumn function, which lower-cases "slicer").
//   2. TableGroup set to "Dimension" instead of "Dim Calendar".
//   3. Mark-as-date-table step is merged in directly (no Selected.Table / reload), no longer creates
//      the "Max Calendar date" / "Last fact table date" measures, and no longer creates a relationship.
//      It hard-codes the "Date" column as the date-table key instead of prompting for a selection.
//   4. Adds two calendar options: "Gregorian" (standard calendar) and "Weekly" (ISO week-based).
//   5. Each data column gets a Description, and SortByColumn is set on the slicer/month/day-name
//      columns (e.g. "Month name" sorted by "Month no", "Week slicer" sorted by "ISO year week").
//   6. Calendar options no longer include "Day no of Month"/"Day no of Week" as time units (the
//      columns still exist on the table, just not exposed via either calendar option). The "Weekly"
//      calendar option's Week time unit now also carries "Week slicer" as an associated column, and
//      its Date time unit now also carries "Date slicer"/"Date_Key" as associated columns (matching
//      the Gregorian calendar's Date time unit).
//
// Requires compatibility level 1702 or above (for Calendar / AddCalendar support).
//
// TE2 COMPATIBILITY NOTE (see te2-script-conversion skill for details): this is a converted copy of
// the TE3 original ("1 - Create English Calendar (combined).cs"); do not hand-edit the two in
// parallel without re-syncing. Changes made for TE2's default (non-Roslyn, ~C#5) compiler:
//   1. Added an explicit #r "System.Drawing" reference - TE2 doesn't auto-reference it like TE3 does,
//      and the file uses `Point`/`Size` for dialog layout.
//   2. Removed the redundant `using System.Linq;` line - already auto-imported by TE2's default
//      script host.
//   3. Rewrote all 33 `$"..."`/`$@"..."` interpolated strings as plain string concatenation (string
//      interpolation is a C# 6 feature not supported by TE2's compiler).
//   4. Rewrote the single type-pattern-matching `is Column siblingCol` (C# 7) as a classic
//      `is Column` check + explicit cast.
//   5. Changed `Model.AddTable(_calendarTableName, false)` (TE3's 2-arg overload) to the single-arg
//      `Model.AddTable(_calendarTableName)` (TE2 doesn't have the 2-arg overload; behavior is the
//      same since a partition is added explicitly afterward either way).
//   6. Removed the `= null` auto-property initializer on `SelectionForm.SelectedOption` (C# 6 feature)
//      - the property already defaults to null without it.
//   7. Removed the redundant `using System;` line - already auto-imported by TE2's default script host
//      (was causing a CS0105 warning).
//   8. Changed `_Table.TableGroup = "Dimension";` to `_Table.SetAnnotation("TabularEditor_TableGroup",
//      "Dimension");` - TE2's older TOM wrapper doesn't expose the TableGroup convenience property;
//      "Table Group" is stored via the "TabularEditor_TableGroup" annotation in both versions, TE3
//      just also surfaces it as a property.
// NOTE: AddCalendar/AddTimeUnit (Calendar Options) calls were left unchanged - not yet confirmed to
// work against TE2's TOM wrapper; if TE2 reports these as missing, Calendar Options may need to be
// skipped/simplified for the TE2 version (ask the user how to proceed if that error surfaces).

#r "Microsoft.VisualBasic"
#r "System.Drawing"
using Microsoft.VisualBasic;
using System.Drawing;
using System.Windows.Forms;

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// ============================================================================
// STEP 1: Create the English Calendar table
// ============================================================================

string _calendarTableName = Interaction.InputBox("Provide the name of the calendar table", "Table Name", "Calendar");

// Check if the table already exists
if (Model.Tables.Contains(_calendarTableName))
{
    Info("Table '" + (_calendarTableName) + "' already exists. Skipping creation.");
    return;
}

// Create the calendar table
var _Table = Model.AddTable(_calendarTableName);

// Define the M expression for partition
string _MPartition = @"
let
    Source = List.Dates(#""Start date"", Number.From(#""End date""- #""Start date"")+1, #duration(1, 0, 0, 0)),
    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #""Renamed Columns"" = Table.RenameColumns(#""Converted to Table"",{{""Column1"", ""Date""}}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Renamed Columns"",{{""Date"", type date}}),

    // Adding calculated columns
    AddedColumns = Table.AddColumn(#""Changed Type"", ""Date_Key"", each Date.Year([Date]) * 10000 + Date.Month([Date]) * 100 + Date.Day([Date]), Int64.Type),

    AddedDateSlicer = Table.AddColumn(AddedColumns, ""Date_slicer"", each
        if [Date] = Date.From(DateTime.LocalNow()) then ""Today""
        else if [Date] = Date.AddDays(Date.From(DateTime.LocalNow()), -1) then ""Yesterday""
        else Date.ToText([Date], ""yyyy-MM-dd""), type text),

    AddedYear = Table.AddColumn(AddedDateSlicer, ""Year"", each Date.Year([Date]), Int64.Type),

    AddedYearSlicer = Table.AddColumn(AddedYear, ""Year_slicer"", each
        if Date.Year([Date]) = Date.Year(DateTime.LocalNow()) then ""Current Year""
        else Text.From(Date.Year([Date])), type text),

    AddedQuarterNo = Table.AddColumn(AddedYearSlicer, ""Quarter_no"", each Date.QuarterOfYear([Date]), Int64.Type),

    AddedQuarter = Table.AddColumn(AddedQuarterNo, ""Quarter"", each ""Q"" & Text.From(Date.QuarterOfYear([Date])), type text),

    AddedYearQuarter = Table.AddColumn(AddedQuarter, ""Year_quarter"", each
        Text.From(Date.Year([Date])) & "" Q"" & Text.From(Date.QuarterOfYear([Date])), type text),

    AddedMonthNo = Table.AddColumn(AddedYearQuarter, ""Month_no"", each Date.Month([Date]), Int64.Type),

    AddedMonthName = Table.AddColumn(AddedMonthNo, ""Month_name"", each Date.ToText([Date], ""MMMM""), type text),

    AddedMonthShortName = Table.AddColumn(AddedMonthName, ""Month_short_name"", each Text.Start(Date.ToText([Date], ""MMMM""), 3), type text),

    AddedYearMonth = Table.AddColumn(AddedMonthShortName, ""Year_month"", each Date.ToText([Date], ""yyyy-MM""), type text),
    AddedYearMonthNo = Table.AddColumn(AddedYearMonth, ""Year_month_no"", each Date.Year([Date])*100+Date.Month([Date]), Int64.Type),
    AddedMonthNameYear = Table.AddColumn(AddedYearMonthNo, ""Month_name_Year"", each Text.Combine({Text.Start(Date.ToText([Date], ""MMMM""), 3), Text.From(Date.Year([Date]), ""en-US"")}, ""-""), type text),
    AddedYearMonthSlicer = Table.AddColumn(AddedMonthNameYear, ""Year_month_slicer"", each if Date.Year([Date]) = Date.Year(DateTime.LocalNow()) and Date.Month([Date]) = Date.Month(DateTime.LocalNow()) then ""Current month"" else [Year_month], type text),
    #""Inserted Day"" = Table.AddColumn(AddedYearMonthSlicer, ""Day_no_of_Month"", each Date.Day([Date]), Int64.Type),
    #""Inserted Day of Week"" = Table.AddColumn(#""Inserted Day"", ""Day_no_of_Week"", each Date.DayOfWeek([Date],Day.Monday), Int64.Type),
    #""Inserted Day Name"" = Table.AddColumn(#""Inserted Day of Week"", ""Day_name"", each Date.DayOfWeekName([Date]), type text),
    #""Inserted is_Weekend"" = Table.AddColumn(#""Inserted Day Name"", ""is_Weekend"", each if Date.DayOfWeek([Date]) >= 5 then 1 else 0, Int64.Type),

    #""Inserted is History"" = Table.AddColumn(#""Inserted is_Weekend"", ""is_History"", each if [Date] < Date.From(DateTime.LocalNow()) then 1 else 0, Int64.Type),
    #""Added ISO_week"" = Table.AddColumn(#""Inserted is History"", ""ISO_week"", each let
   CurrentThursday =  Date.AddDays( [Date], 3 - Date.DayOfWeek( [Date], Day.Monday ) ),
   YearCurrThursday = Date.Year( CurrentThursday ),
   FirstThursdayOfYear =
         Date.AddDays( #date(YearCurrThursday, 1, 7) ,
                       - Date.DayOfWeek( #date( YearCurrThursday, 1, 1 ), Day.Friday ) ),
   ISO_Week = Duration.Days( CurrentThursday - FirstThursdayOfYear ) / 7 + 1
 in
   ISO_Week, Int64.Type),
    #""Added ISO_year"" = Table.AddColumn(#""Added ISO_week"", ""ISO_year"", each Date.Year(
       Date.AddDays( [Date], 26 - [ISO_week] ) ), Int64.Type),
    #""Added ISO year week"" = Table.AddColumn(#""Added ISO_year"", ""ISO_year_week"", each let

    ISO_Year = [ISO_year],
    ISO_Week = [ISO_week],

    // Format the week with leading zero if necessary
    ISO_WeekFormatted = Text.PadStart(Text.From(ISO_Week), 2, ""0""),

    // Combine ISO_Year and the formatted ISO_Week
    ISO_Year_Week = Text.Combine({"""", Text.From(ISO_Year), ""-W"", ISO_WeekFormatted})
in
    ISO_Year_Week, type text),
    AddWeekSlicer = Table.AddColumn(#""Added ISO year week"", ""Week_slicer"", each
        if [ISO_year] = Date.Year(DateTime.LocalNow()) and [ISO_week] = Date.WeekOfYear(DateTime.LocalNow(), Day.Monday)
        then ""Current_Week""
        else Text.From([ISO_year_week]),
        type text),
    Calendar = Table.TransformColumnNames(AddWeekSlicer, each RenameColumn(_))
in
    Calendar";

try
{
    // Add partition
    var partition = _Table.AddMPartition(_calendarTableName, _MPartition);
    partition.Mode = ModeType.Import; // Set refresh policy to Import

    // Adding data columns
    _Table.AddDataColumn("Date", "Date", "6. Date", DataType.DateTime)
        .Description = "The calendar date.";
    var _key = _Table.AddDataColumn("Date_Key", "Date_Key", "6. Date", DataType.Int64);
    _key.Description = "Integer surrogate key for the date, in YYYYMMDD format (e.g. 20250131).";

    _Table.AddDataColumn("Year", "Year", "2. Year", DataType.Int64)
        .Description = "The calendar year number (e.g. 2025).";

    _Table.AddDataColumn("Month no", "Month no", "4. Month", DataType.Int64)
        .Description = "The month number within the year (1-12), used to sort 'Month name'/'Month short name'.";
    _Table.AddDataColumn("Month name", "Month name", "4. Month", DataType.String)
        .Description = "The full month name (e.g. 'January').";
    _Table.AddDataColumn("Month short name", "Month short name", "4. Month", DataType.String)
        .Description = "The 3-letter abbreviated month name (e.g. 'Jan').";
    _Table.AddDataColumn("Year month", "Year month", "4. Month", DataType.String)
        .Description = "The year and month as text, in 'yyyy-MM' format (e.g. '2025-01').";
    _Table.AddDataColumn("Year month no", "Year month no", "4. Month", DataType.Int64)
        .Description = "The year and month as a sortable integer, in YYYYMM format (e.g. 202501), used to sort 'Year month'/'Month name Year'/'Year month slicer'.";
    _Table.AddDataColumn("Month name Year", "Month name Year", "4. Month", DataType.String)
        .Description = "The abbreviated month name and year (e.g. 'Jan-2025').";

    _Table.AddDataColumn("Quarter", "Quarter", "3. Quarter", DataType.String)
        .Description = "The quarter of the year, as text (e.g. 'Q1'), used to sort 'Year quarter'.";
    _Table.AddDataColumn("Quarter no", "Quarter no", "3. Quarter", DataType.Int64)
        .Description = "The quarter number within the year (1-4).";
    _Table.AddDataColumn("Year quarter", "Year quarter", "3. Quarter", DataType.String)
        .Description = "The year and quarter as text (e.g. '2025 Q1').";

    _Table.AddDataColumn("ISO week", "ISO week", "5. Week", DataType.Int64)
        .Description = "The ISO 8601 week number within the ISO year (1-53).";
    _Table.AddDataColumn("ISO year", "ISO year", "5. Week", DataType.Int64)
        .Description = "The ISO 8601 year (can differ from the calendar year for dates near year-end/-start).";
    _Table.AddDataColumn("ISO year week", "ISO year week", "5. Week", DataType.String)
        .Description = "The ISO 8601 year and week as text, in 'yyyy-Www' format (e.g. '2025-W05'), used to sort 'Week slicer'.";

    _Table.AddDataColumn("Day name", "Day name", "7. Day", DataType.String)
        .Description = "The day of week name (e.g. 'Monday').";
    _Table.AddDataColumn("Day no of Week", "Day no of Week", "7. Day", DataType.Int64)
        .Description = "The day number within the week (1 = Monday .. 7 = Sunday), used to sort 'Day name'.";
    _Table.AddDataColumn("Day no of Month", "Day no of Month", "7. Day", DataType.Int64)
        .Description = "The day number within the month (1-31).";

    // Change #1: "Year Slicer" -> "Year slicer" (matches the M-generated column name)
    _Table.AddDataColumn("Year slicer", "Year slicer", "1. Slicer", DataType.String)
        .Description = "'Current Year' for the current year, otherwise the year number as text - for use in slicers.";
    _Table.AddDataColumn("Year month slicer", "Year month slicer", "1. Slicer", DataType.String)
        .Description = "'Current month' for the current year/month, otherwise 'Year month' - for use in slicers.";
    _Table.AddDataColumn("Week slicer", "Week slicer", "1. Slicer", DataType.String)
        .Description = "'Current_Week' for the current ISO year/week, otherwise 'ISO year week' - for use in slicers.";
    _Table.AddDataColumn("Date slicer", "Date slicer", "1. Slicer", DataType.String)
        .Description = "'Today'/'Yesterday' for those dates, otherwise the date as 'yyyy-MM-dd' text - for use in slicers.";

    _Table.AddDataColumn("is_History", "is_History", "8. Boolean", DataType.Int64)
        .Description = "1 if the date is before today, otherwise 0.";
    _Table.AddDataColumn("is_Weekend", "is_Weekend", "8. Boolean", DataType.Int64)
        .Description = "1 if the date falls on a Saturday or Sunday, otherwise 0.";

    _key.IsHidden = true;
    _key.IsAvailableInMDX = false;

    Model.Tables[_calendarTableName].DataCategory = "Time";

    // ---- Sort by (slicers, months, day names) ----
    _Table.Columns["Month name"].SortByColumn      = _Table.Columns["Month no"];
    _Table.Columns["Month short name"].SortByColumn = _Table.Columns["Month no"];
    _Table.Columns["Month name Year"].SortByColumn = _Table.Columns["Year month no"];
    _Table.Columns["Year month"].SortByColumn      = _Table.Columns["Year month no"];
    _Table.Columns["Day name"].SortByColumn        = _Table.Columns["Day no of Week"];

    _Table.Columns["Year slicer"].SortByColumn       = _Table.Columns["Year"];
    _Table.Columns["Year month slicer"].SortByColumn = _Table.Columns["Year month no"];
    _Table.Columns["Week slicer"].SortByColumn       = _Table.Columns["ISO year week"];
    _Table.Columns["Date slicer"].SortByColumn       = _Table.Columns["Date"];
}
catch (Exception ex)
{
    Info("Failed to create Date-table.\n\nReason: " + (ex.Message));
    return;
}

// Function to create a form with radio buttons
public class SelectionForm : Form
{
    public string SelectedOption { get; private set; }
    private RadioButton rbRelative;
    private RadioButton rbSelected;
    private Button btnOK;

    public SelectionForm()
    {
        this.Text = "Select Date Type";
        this.Size = new Size(400, 220);
        this.StartPosition = FormStartPosition.CenterScreen;

        Label lblPrompt = new Label()
        {
            Text = "Choose an option:",
            Location = new Point(20, 20),
            AutoSize = true
        };

        rbRelative = new RadioButton()
        {
            Text = "Relative Date",
            Location = new Point(20, 50),
            AutoSize = true
        };

        rbSelected = new RadioButton()
        {
            Text = "Selected Date",
            Location = new Point(20, 80),
            AutoSize = true
        };

        btnOK = new Button()
        {
            Text = "OK",
            Location = new Point(100, 120),
            Size = new Size(75, 30),
            DialogResult = DialogResult.OK
        };

        btnOK.Click += (sender, e) =>
        {
            if (rbRelative.Checked)
                SelectedOption = "Relative date";
            else if (rbSelected.Checked)
                SelectedOption = "Selected date";
            else
            {
                MessageBox.Show("Please select an option before proceeding.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;  // Do not close the form
            }
            this.Close();
        };

        this.Controls.Add(lblPrompt);
        this.Controls.Add(rbRelative);
        this.Controls.Add(rbSelected);
        this.Controls.Add(btnOK);
    }
}

// Show selection form
SelectionForm form = new SelectionForm();
form.ShowDialog();

// Get user selection
string selectedOption = form.SelectedOption;

// Exit script if no selection was made
if (string.IsNullOrEmpty(selectedOption))
{
    Info("No option selected. Script aborted.");
    return;
}

if (selectedOption == "Relative date")
{
    // Prompt for start date input
    string inputStart = Interaction.InputBox("Select number of historical years:", "No of years", "-10");
    int _yearsFromStart;
    if (!int.TryParse(inputStart, out _yearsFromStart))
    {
        Info("Invalid input. Using default value: -10");
        _yearsFromStart = -10;
    }

    // Prompt for end date input
    string inputEnd = Interaction.InputBox("Select number of years from current year:", "No of years", "0");
    int _yearsFromEnd;
    if (!int.TryParse(inputEnd, out _yearsFromEnd))
    {
        Info("Invalid input. Using default value: 0");
        _yearsFromEnd = 0;
    }

    // Create Relative StartDate
    string expressionStartName = "Start date";
    if (!Model.Expressions.Contains(expressionStartName))
    {
        string _MExpressionStart = "\nlet\n    Source = Date.StartOfYear(Date.AddYears(DateTime.LocalNow(), " + (_yearsFromStart) + ")),\n    #\"Extracted Date\" = Date.From(Source)\nin\n    #\"Extracted Date\"\n";
        try
        {
            var sharedExpression = Model.AddExpression(expressionStartName, _MExpressionStart);
            sharedExpression.SetAnnotation("Start date", "Import");
            Info("Created Shared Expression '" + (expressionStartName) + "' successfully.");
        }
        catch (Exception ex)
        {
            Info("Failed to create Shared Expression '" + (expressionStartName) + "'.\n\nReason: " + (ex.Message));
        }
    }
    else
    {
        Info("Skipped creation: Shared Expression '" + (expressionStartName) + "' already exists.");
    }

    // Create Relative EndDate
    string expressionEndName = "End date";
    if (!Model.Expressions.Contains(expressionEndName))
    {
        string _MExpressionEnd = "\nlet\n    Source = Date.EndOfYear(Date.AddYears(DateTime.LocalNow(), " + (_yearsFromEnd) + ")),\n    #\"Extracted Date\" = Date.From(Source)\nin\n    #\"Extracted Date\"\n";
        try
        {
            var sharedExpression = Model.AddExpression(expressionEndName, _MExpressionEnd);
            sharedExpression.SetAnnotation("End date", "Import");
            Info("Created Shared Expression '" + (expressionEndName) + "' successfully.");
        }
        catch (Exception ex)
        {
            Info("Failed to create Shared Expression '" + (expressionEndName) + "'.\n\nReason: " + (ex.Message));
        }
    }
    else
    {
        Info("Skipped creation: Shared Expression '" + (expressionEndName) + "' already exists.");
    }
}
else if (selectedOption == "Selected date")
{
    var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime).ToList();

    // Select Earliest Date Column
    string _EarliestDate = SelectColumn(_dateColumns, null, "Select the Column with the Earliest Date:")
        .DaxObjectFullName;

    // Extract Table Name and Column Name
    var tableNameEarliest = _EarliestDate.Split('[')[0].Trim('\'');  // Remove single quotes around table name
    var columnNameEarliest = _EarliestDate.Substring(_EarliestDate.IndexOf('[')); // Keep column reference as is

    // Format table name
    string _formattedEarliestDate = "#\"" + (tableNameEarliest) + "\"" + (columnNameEarliest);

    string expressionEarliestName = "Start date";

    if (!Model.Expressions.Contains(expressionEarliestName))
    {
        string _MExpressionEarliest = "\nlet\n    Source = List.Min(" + (_formattedEarliestDate) + "),\n    #\"Extracted Date\" = Date.From(Source),\n    #\"Calculated Start of Year\" = Date.StartOfYear(#\"Extracted Date\")\nin\n    #\"Calculated Start of Year\"\n";
        try
        {
            var sharedExpression = Model.AddExpression(expressionEarliestName, _MExpressionEarliest);
            sharedExpression.SetAnnotation("Start date", "Import");
            Info("Created Shared Expression '" + (expressionEarliestName) + "' successfully.");
        }
        catch (Exception ex)
        {
            Info("Failed to create Shared Expression '" + (expressionEarliestName) + "'.\n\nReason: " + (ex.Message));
        }
    }
    else
    {
        Info("Skipped creation: Shared Expression '" + (expressionEarliestName) + "' already exists.");
    }

    // Select Latest Date Column
    string _LatestDate = SelectColumn(_dateColumns, null, "Select the Column with the Latest Date:")
        .DaxObjectFullName;

    // Extract Table Name and Column Name
    var tableNameLatest = _LatestDate.Split('[')[0].Trim('\'');  // Remove single quotes around table name
    var columnNameLatest = _LatestDate.Substring(_LatestDate.IndexOf('[')); // Keep column reference as is

    // Format table name
    string _formattedLatestDate = "#\"" + (tableNameLatest) + "\"" + (columnNameLatest);

    string expressionLatestName = "End date";

    if (!Model.Expressions.Contains(expressionLatestName))
    {
        string _MExpressionLatest = "\nlet\n    Source = List.Max(" + (_formattedLatestDate) + "),\n    #\"Extracted Date\" = Date.From(Source),\n    #\"Calculated End of Year\" = Date.EndOfYear(#\"Extracted Date\")\nin\n    #\"Calculated End of Year\"\n";
        try
        {
            var sharedExpression = Model.AddExpression(expressionLatestName, _MExpressionLatest);
            sharedExpression.SetAnnotation("End date", "Import");
            Info("Created Shared Expression '" + (expressionLatestName) + "' successfully.");
        }
        catch (Exception ex)
        {
            Info("Failed to create Shared Expression '" + (expressionLatestName) + "'.\n\nReason: " + (ex.Message));
        }
    }
    else
    {
        Info("Skipped creation: Shared Expression '" + (expressionLatestName) + "' already exists.");
    }
}
else
{
    Info("No option selected. Script aborted.");
}

// Insert Rename_columns, as input is using _ if the same can be used in Lakehouse

// Function name
string functionName = "RenameColumn";

// Check if the function already exists
if (Model.Expressions.Contains(functionName))
{
    Info("Function '" + (functionName) + "' already exists. Skipping creation.");
}
else
{
    // Define the M function expression
    string _MFunction = @"
(Column as text) as text =>
let
    Source =
        if     Text.StartsWith(Column, ""PK"") then Column
        else if Text.StartsWith(Column, ""FK"") then Column
        else if Text.StartsWith(Column, ""is_"") then Column
        else if Text.StartsWith(Column, ""er_"") then Column
        else if Text.EndsWith  (Column, ""Id"") then Column
        else if Text.EndsWith  (Column, ""Key"") then Column
        else  Text.Combine(Splitter.SplitTextByCharacterTransition({""a""..""z""}, {""A""..""Z""})(Text.Replace(Column, ""_"", "" "")), "" "")
    in
        Source";

    // Add the M function to the model
    var mFunction = Model.AddExpression(functionName, _MFunction);

    Info("Created function '" + (functionName) + "' successfully.");
}

// ============================================================================
// STEP 2: Mark the table as a Date table
// (merged in directly - no reload, no measures, no relationship, hard-coded to "Date")
// ============================================================================

try
{
    string _dateColumnName = "Date";

    if (!_Table.Columns.Contains(_dateColumnName))
    {
        Error("Column '" + (_dateColumnName) + "' not found in table '" + (_calendarTableName) + "'. Cannot mark as date table.");
    }
    else
    {
        _Table.DataCategory = "Time";
        _Table.Columns[_dateColumnName].IsKey = true;

        // Change #2: Table group should be "Dimension"
        // TE2's older TOM wrapper doesn't expose a TableGroup property directly - "Table Group" is
        // implemented via the "TabularEditor_TableGroup" annotation (TE3 surfaces the same value as a
        // convenience property called TableGroup).
        _Table.SetAnnotation("TabularEditor_TableGroup", "Dimension");

        // Best practice: prevent numeric columns from being summarized
        foreach (var column in _Table.Columns)
        {
            if (column.DataType == DataType.Int64)
                column.SummarizeBy = AggregateFunction.None;
        }

        // Best practice: apply a consistent format string to date columns
        foreach (var column in _Table.Columns)
        {
            if (column.DataType == DataType.DateTime)
                column.FormatString = "MM/dd/yyyy";
        }

        Info("'" + (_calendarTableName) + "' marked as date table using column '" + (_dateColumnName) + "'.");
    }
}
catch (Exception ex)
{
    Error("Error marking '" + (_calendarTableName) + "' as date table: " + (ex.Message));
}

// ============================================================================
// STEP 3: Add calendar options (Gregorian + Weekly)
// ============================================================================

try
{
    // Local helper to safely fetch a column by name
    Func<string, Column> col = (name) => _Table.Columns.Contains(name) ? _Table.Columns[name] : null;

    // ----- Gregorian calendar (standard calendar) -----
    var calendar1 = _Table.AddCalendar("Gregorian");
    calendar1.Description = "Standard calendar using calendar option";

    calendar1.AddTimeUnit(TimeUnit.Year, col("Year"), col("Year slicer"));
    calendar1.AddTimeUnit(TimeUnit.Quarter, col("Year quarter"));
    calendar1.AddTimeUnit(TimeUnit.QuarterOfYear, col("Quarter no"), col("Quarter"));
    calendar1.AddTimeUnit(TimeUnit.Month, col("Year month no"), col("Year month"), col("Year month slicer"), col("Month name Year"));
    calendar1.AddTimeUnit(TimeUnit.MonthOfYear, col("Month no"), col("Month name"), col("Month short name"));
    calendar1.AddTimeUnit(TimeUnit.Date, col("Date"), col("Date slicer"), col("Date_Key"));

    Info("Created calendar option 'Gregorian'.");

    // ----- Weekly calendar (ISO week-based) -----
    var calendar2 = _Table.AddCalendar("Weekly");
    calendar2.Description = "Week ";

    calendar2.AddTimeUnit(TimeUnit.Year, col("ISO year"));
    calendar2.AddTimeUnit(TimeUnit.Week, col("ISO year week"), col("Week slicer"));
    calendar2.AddTimeUnit(TimeUnit.WeekOfYear, col("ISO week"));
    calendar2.AddTimeUnit(TimeUnit.Date, col("Date"), col("Date slicer"), col("Date_Key"));

    Info("Created calendar option 'Weekly'.");
}
catch (Exception ex)
{
    Error("Error creating calendar options: " + (ex.Message));
}

// ============================================================================
// STEP 4: Relationships to other tables (optional)
// ============================================================================
// Dialog A - pick a mode: skip, connect via the "Date" (DateTime) column, or connect via the
// "Date_Key" (Int64) surrogate key column. Either mode reuses the same table->column picker (Dialog B)
// for the ACTIVE relationship (max one column per table) and, optionally, a second pass (Dialog C) to
// add further INACTIVE relationships (multiple columns per table allowed).
string _relationshipMode = "Skip";
{
    var modeForm = new Form { Text = "Relationships to other tables", Width = 480, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    var modeFlow = new FlowLayoutPanel { Padding = new Padding(15), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    modeForm.Controls.Add(modeFlow);

    modeFlow.Controls.Add(new Label { Text = "Create relationships from other tables to this calendar table?", AutoSize = true, Margin = new Padding(0, 0, 0, 10) });
    var rbSkip = new RadioButton { Text = "Skip relationships", AutoSize = true, Checked = true };
    var rbDate = new RadioButton { Text = "Connect with Date-format (uses '" + (_calendarTableName) + "'[Date])", AutoSize = true };
    var rbInt  = new RadioButton { Text = "Connect with Int-format (uses '" + (_calendarTableName) + "'[Date_Key])", AutoSize = true };
    foreach (var rb in new[] { rbSkip, rbDate, rbInt }) modeFlow.Controls.Add(rb);

    var modeBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
    var modeOk = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    var modeCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
    modeBtnPanel.Controls.Add(modeOk);
    modeBtnPanel.Controls.Add(modeCancel);
    modeFlow.Controls.Add(modeBtnPanel);
    modeForm.AcceptButton = modeOk;
    modeForm.CancelButton = modeCancel;

    if (modeForm.ShowDialog() != DialogResult.OK)
        _relationshipMode = "Cancel";
    else if (rbDate.Checked)
        _relationshipMode = "Date";
    else if (rbInt.Checked)
        _relationshipMode = "Int";
}

if (_relationshipMode == "Cancel")
{
    Info("Cancelled - no relationships were created.");
}
else if (_relationshipMode == "Skip")
{
    Info("Skipped relationship creation, as requested.");
}
else
{
    // Generic table->column picker: a two-level TreeView (Table -> candidate Column) with checkboxes
    // and a live search box filtering by table or column name, similar in spirit to the namespace-tree
    // picker in "5 - Create measure from model dependent functions.cs". When singleSelectPerTable is
    // true, checking a column automatically unchecks any other checked column under the same table
    // (max one selection per table); table (parent) nodes themselves are not selectable.
    Func<List<Column>, bool, string, string, List<Column>> PickColumnsDialog = (candidates, singleSelectPerTable, dialogTitle, instructions) =>
    {
        var picked = new List<Column>();
        if (candidates == null || candidates.Count == 0) return picked;

        var pickForm = new Form { Text = dialogTitle, Width = 560, Height = 620, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        var pickLabel = new Label { Text = instructions, Location = new Point(15, 15), Width = 515, Height = 45 };
        var searchLabel = new Label { Text = "Search:", Location = new Point(15, 65), AutoSize = true };
        var searchBox = new TextBox { Location = new Point(70, 62), Width = 460 };
        var pickTree = new TreeView { Location = new Point(15, 90), Width = 515, Height = 430, CheckBoxes = true };
        var pickOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(365, 535), AutoSize = true };
        var pickCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(450, 535), AutoSize = true };
        pickForm.Controls.Add(pickLabel);
        pickForm.Controls.Add(searchLabel);
        pickForm.Controls.Add(searchBox);
        pickForm.Controls.Add(pickTree);
        pickForm.Controls.Add(pickOk);
        pickForm.Controls.Add(pickCancel);
        pickForm.AcceptButton = pickOk;
        pickForm.CancelButton = pickCancel;

        // Column.DaxObjectFullName -> checked state, kept across search-filter rebuilds of the tree.
        var checkedNames = new HashSet<string>();
        bool suppress = false;

        Action<string> Rebuild = (filterText) =>
        {
            suppress = true;
            pickTree.BeginUpdate();
            pickTree.Nodes.Clear();

            var filtered = candidates
                .Where(c => string.IsNullOrWhiteSpace(filterText)
                    || c.Table.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
                    || c.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(c => c.Table.Name).ThenBy(c => c.Name)
                .ToList();

            var tableNodes = new Dictionary<string, TreeNode>();
            foreach (var c in filtered)
            {
                TreeNode tableNode;
                if (!tableNodes.TryGetValue(c.Table.Name, out tableNode))
                {
                    tableNode = new TreeNode(c.Table.Name);
                    pickTree.Nodes.Add(tableNode);
                    tableNodes[c.Table.Name] = tableNode;
                }
                var leaf = new TreeNode(c.Name) { Tag = c, Checked = checkedNames.Contains(c.DaxObjectFullName) };
                tableNode.Nodes.Add(leaf);
            }

            pickTree.ExpandAll();
            pickTree.EndUpdate();
            suppress = false;
        };

        Rebuild("");
        searchBox.TextChanged += (s, e) => Rebuild(searchBox.Text);

        // Table (parent) nodes have no Tag and are not selectable.
        pickTree.BeforeCheck += (s, e) => { if (!suppress && e.Node.Tag == null) e.Cancel = true; };

        pickTree.AfterCheck += (s, e) =>
        {
            if (suppress || e.Node.Tag == null) return;
            var thisCol = (Column)e.Node.Tag;

            suppress = true;
            if (e.Node.Checked)
            {
                checkedNames.Add(thisCol.DaxObjectFullName);
                if (singleSelectPerTable)
                {
                    foreach (TreeNode sibling in e.Node.Parent.Nodes)
                    {
                        if (sibling != e.Node && sibling.Checked)
                        {
                            sibling.Checked = false;
                            if (sibling.Tag is Column) { var siblingCol = (Column)sibling.Tag; checkedNames.Remove(siblingCol.DaxObjectFullName); }
                        }
                    }
                }
            }
            else
            {
                checkedNames.Remove(thisCol.DaxObjectFullName);
            }
            suppress = false;
        };

        if (pickForm.ShowDialog() != DialogResult.OK) return picked;

        picked = candidates.Where(c => checkedNames.Contains(c.DaxObjectFullName)).ToList();
        return picked;
    };

    string oneSideColumnName = _relationshipMode == "Date" ? "Date" : "Date_Key";
    var oneSideColumn = _Table.Columns[oneSideColumnName];
    DataType candidateType = _relationshipMode == "Date" ? DataType.DateTime : DataType.Int64;
    string typeLabel = _relationshipMode == "Date" ? "date/time" : "whole number (Int64)";

    var candidateColumns = Model.AllColumns
        .Where(c => c.Table.Name != _calendarTableName && c.DataType == candidateType)
        .ToList();

    // For the Int-format mode only, narrow the candidate tables down to those that already appear as
    // the MANY side of an existing relationship elsewhere in the model - a simple heuristic to prefer
    // likely fact tables over dimension tables that merely happen to have an Int64 column. Not applied
    // for Date-format, since date/time columns are typically only found on fact tables anyway.
    if (_relationshipMode == "Int")
    {
        var manySideTableNames = Model.Relationships.Select(r => r.FromTable.Name).Distinct().ToList();
        if (manySideTableNames.Any())
            candidateColumns = candidateColumns.Where(c => manySideTableNames.Contains(c.Table.Name)).ToList();
    }

    if (candidateColumns.Count == 0)
    {
        Info("No " + (typeLabel) + " columns found on other tables - skipping relationship creation.");
    }
    else
    {
        Func<Column, Column, bool> RelationshipExists = (fromCol, toCol) =>
            Model.Relationships.Any(r => r.FromColumn == fromCol && r.ToColumn == toCol);

        var activeSelections = PickColumnsDialog(candidateColumns, true,
            "Select the ACTIVE relationship column",
            "For each table below, optionally check ONE " + (typeLabel) + " column to create an ACTIVE relationship to '" + (_calendarTableName) + "'[" + (oneSideColumnName) + "]. Leave a table unchecked to skip it.");

        foreach (var fromCol in activeSelections)
        {
            if (RelationshipExists(fromCol, oneSideColumn))
            {
                Info("Relationship from '" + (fromCol.DaxObjectFullName) + "' to '" + (oneSideColumn.DaxObjectFullName) + "' already exists - skipping.");
                continue;
            }
            var rel = Model.AddRelationship();
            rel.FromColumn = fromCol;
            rel.ToColumn = oneSideColumn;
            rel.FromCardinality = RelationshipEndCardinality.Many;
            rel.ToCardinality = RelationshipEndCardinality.One;
            rel.IsActive = true;
        }

        var addInactive = MessageBox.Show(
            "Do you want to add further INACTIVE relationships as well (e.g. multiple date roles per fact table)?" + Environment.NewLine +
            "Click 'Yes' to select additional columns, or 'No' to finish.",
            "Additional inactive relationships", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (addInactive == DialogResult.Yes)
        {
            // Exclude columns that already have a relationship to the one-side column (e.g. the one
            // just made active above), so they aren't offered again for an inactive relationship.
            var alreadyRelated = new HashSet<string>(
                Model.Relationships.Where(r => r.ToColumn == oneSideColumn).Select(r => r.FromColumn.DaxObjectFullName));
            var inactiveCandidates = candidateColumns.Where(c => !alreadyRelated.Contains(c.DaxObjectFullName)).ToList();

            if (inactiveCandidates.Count == 0)
            {
                Info("No further candidate columns available for additional inactive relationships.");
            }
            else
            {
                var inactiveSelections = PickColumnsDialog(inactiveCandidates, false,
                    "Select INACTIVE relationship columns",
                    "Select any number of additional " + (typeLabel) + " columns (multiple per table allowed) to create as INACTIVE relationships to '" + (_calendarTableName) + "'[" + (oneSideColumnName) + "].");

                foreach (var fromCol in inactiveSelections)
                {
                    if (RelationshipExists(fromCol, oneSideColumn))
                    {
                        Info("Relationship from '" + (fromCol.DaxObjectFullName) + "' to '" + (oneSideColumn.DaxObjectFullName) + "' already exists - skipping.");
                        continue;
                    }
                    var rel = Model.AddRelationship();
                    rel.FromColumn = fromCol;
                    rel.ToColumn = oneSideColumn;
                    rel.FromCardinality = RelationshipEndCardinality.Many;
                    rel.ToCardinality = RelationshipEndCardinality.One;
                    rel.IsActive = false;
                }
            }
        }

        Info("Relationship setup completed.");
    }
}

Info("Calendar table creation, date-table setup, calendar options, and relationships completed successfully!");
