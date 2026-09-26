// The purpose of this macro is to create a Norwegian calendar table, mark it as a date table,
// and configure calendar options (Gregoriansk + ISO/Ukentlig) - all in a single script execution
// (no reload required between steps, since everything runs against the same in-memory model).
//
// This is the Norwegian-language equivalent of "1 - Create English Calendar (combined).cs":
// same structure and logic, but the table/column names, descriptions and calendar option names
// are in Norwegian (matching the naming already used in "1 - Lage Norsk kalender.cs"). Dialogs,
// info boxes and code comments remain in English.
//
// In addition, this script adds an English ("en-US") translation of the table name, all column
// names and all column descriptions, using the exact English wording from the combined English
// calendar macro. The "en-US" culture is created automatically if it does not already exist.
//
// Requires compatibility level 1702 or above (for Calendar / AddCalendar support).

#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// ============================================================================
// STEP 1: Create the Norwegian Calendar table
// ============================================================================

string _calendarTableName = Interaction.InputBox("Provide the name of the calendar table", "Table Name", "Kalender");

// Check if the table already exists
if (Model.Tables.Contains(_calendarTableName))
{
    Info($"Table '{_calendarTableName}' already exists. Skipping creation.");
    return;
}

// Create the calendar table
var _Table = Model.AddTable(_calendarTableName, false);

// Define the M expression for partition
string _MPartition = @"
let
    Source = List.Dates(#""Start date"", Number.From(#""End date""- #""Start date"")+1, #duration(1, 0, 0, 0)),
    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #""Renamed Columns"" = Table.RenameColumns(#""Converted to Table"",{{""Column1"", ""Dato""}}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Renamed Columns"",{{""Dato"", type date}}),

    // Adding calculated columns
    AddedColumns = Table.AddColumn(#""Changed Type"", ""Dato_Id"", each Date.Year([Dato]) * 10000 + Date.Month([Dato]) * 100 + Date.Day([Dato]), Int64.Type),

    AddedDateSlicer = Table.AddColumn(AddedColumns, ""Dato_filter"", each
        if [Dato] = Date.From(DateTime.LocalNow()) then ""Today""
        else if [Dato] = Date.AddDays(Date.From(DateTime.LocalNow()), -1) then ""Yesterday""
        else Date.ToText([Dato], ""yyyy-MM-dd""), type text),

    AddedYear = Table.AddColumn(AddedDateSlicer, ""År"", each Date.Year([Dato]), Int64.Type),

    AddedYearSlicer = Table.AddColumn(AddedYear, ""År_filter"", each
        if Date.Year([Dato]) = Date.Year(DateTime.LocalNow()) then ""Inneværende år""
        else Text.From(Date.Year([Dato])), type text),

    AddedQuarterNo = Table.AddColumn(AddedYearSlicer, ""Kvartal_nr"", each Date.QuarterOfYear([Dato]), Int64.Type),

    AddedQuarter = Table.AddColumn(AddedQuarterNo, ""Kvartal"", each ""Q"" & Text.From(Date.QuarterOfYear([Dato])), type text),

    AddedYearQuarter = Table.AddColumn(AddedQuarter, ""År_kvartal"", each
        Text.From(Date.Year([Dato])) & "" Q"" & Text.From(Date.QuarterOfYear([Dato])), type text),

    AddedMonthNo = Table.AddColumn(AddedYearQuarter, ""Måned_nr"", each Date.Month([Dato]), Int64.Type),

    AddedMonthName = Table.AddColumn(AddedMonthNo, ""Måned"", each Text.Proper(Date.ToText([Dato], ""MMMM"", ""nb-NO"")), type text),

    AddedMonthShortName = Table.AddColumn(AddedMonthName, ""Måned_kort_navn"", each Text.Proper(Text.Start(Date.ToText([Dato], ""MMMM"", ""nb-NO""), 3)), type text),

    AddedYearMonth = Table.AddColumn(AddedMonthShortName, ""År_måned"", each Date.ToText([Dato], ""yyyy-MM""), type text),
    AddedYearMonthNo = Table.AddColumn(AddedYearMonth, ""År_måned_no"", each Date.Year([Dato])*100+Date.Month([Dato]), Int64.Type),
    AddedMonthNameYear = Table.AddColumn(AddedYearMonthNo, ""Måned_år"", each Text.Combine({Text.Proper(Text.Start(Date.ToText([Dato], ""MMMM"", ""nb-NO""), 3)), Text.From(Date.Year([Dato]), ""en-US"")}, ""-""), type text),
    AddedYearMonthSlicer = Table.AddColumn(AddedMonthNameYear, ""År_måned_filter"", each if Date.Year([Dato]) = Date.Year(DateTime.LocalNow()) and Date.Month([Dato]) = Date.Month(DateTime.LocalNow()) then ""Denne perioden"" else [År_måned], type text),
    #""Inserted Day"" = Table.AddColumn(AddedYearMonthSlicer, ""Dag_i_måned"", each Date.Day([Dato]), Int64.Type),
    #""Inserted Day of Week"" = Table.AddColumn(#""Inserted Day"", ""Dag_i_uke"", each Date.DayOfWeek([Dato],Day.Monday), Int64.Type),
    #""Inserted Day Name"" = Table.AddColumn(#""Inserted Day of Week"", ""Dag"", each Text.Proper(Date.ToText([Dato], ""dddd"", ""nb-NO"")), type text),
    #""Inserted is_Weekend"" = Table.AddColumn(#""Inserted Day Name"", ""er_Helg"", each if Date.DayOfWeek([Dato]) >= 5 then 1 else 0, Int64.Type),

    #""Inserted is History"" = Table.AddColumn(#""Inserted is_Weekend"", ""er_Historie"", each if [Dato] < Date.From(DateTime.LocalNow()) then 1 else 0, Int64.Type),
    #""Added ISO_week"" = Table.AddColumn(#""Inserted is History"", ""ISO_Uke"", each let
   CurrentThursday =  Date.AddDays( [Dato], 3 - Date.DayOfWeek( [Dato], Day.Monday ) ),
   YearCurrThursday = Date.Year( CurrentThursday ),
   FirstThursdayOfYear =
         Date.AddDays( #date(YearCurrThursday, 1, 7) ,
                       - Date.DayOfWeek( #date( YearCurrThursday, 1, 1 ), Day.Friday ) ),
   ISO_Week = Duration.Days( CurrentThursday - FirstThursdayOfYear ) / 7 + 1
 in
   ISO_Week, Int64.Type),
    #""Added ISO_year"" = Table.AddColumn(#""Added ISO_week"", ""ISO_år"", each Date.Year(
       Date.AddDays( [Dato], 26 - [ISO_Uke] ) ), Int64.Type),
    #""Added ISO year week"" = Table.AddColumn(#""Added ISO_year"", ""ISO_år_uke"", each let

    ISO_Year = [ISO_år],
    ISO_Week = [ISO_Uke],

    // Format the week with leading zero if necessary
    ISO_WeekFormatted = Text.PadStart(Text.From(ISO_Week), 2, ""0""),

    // Combine ISO_Year and the formatted ISO_Week
    ISO_Year_Week = Text.Combine({"""", Text.From(ISO_Year), ""-U"", ISO_WeekFormatted})
in
    ISO_Year_Week, type text),
    AddWeekSlicer = Table.AddColumn(#""Added ISO year week"", ""Uke_filter"", each
        if [ISO_år] = Date.Year(DateTime.LocalNow()) and [ISO_Uke] = Date.WeekOfYear(DateTime.LocalNow(), Day.Monday)
        then ""Denne uken""
        else Text.From([ISO_år_uke]),
        type text),
    Kalender = Table.TransformColumnNames(AddWeekSlicer, each RenameColumn(_))
in
    Kalender";

try
{
    // Add partition
    var partition = _Table.AddMPartition(_calendarTableName, _MPartition);
    partition.Mode = ModeType.Import; // Set refresh policy to Import

    // Adding data columns
    _Table.AddDataColumn("Dato", "Dato", "6. Dato", DataType.DateTime)
        .Description = "Kalenderdatoen.";
    var _key = _Table.AddDataColumn("Dato_Id", "Dato_Id", "6. Dato", DataType.Int64);
    _key.Description = "Heltalls-surrogatnøkkel for datoen, i formatet ÅÅÅÅMMDD (f.eks. 20250131).";

    _Table.AddDataColumn("År", "År", "2. År", DataType.Int64)
        .Description = "Kalenderåret (f.eks. 2025).";

    _Table.AddDataColumn("Måned nr", "Måned nr", "4. Måned", DataType.Int64)
        .Description = "Månedsnummeret innenfor året (1-12), brukt til å sortere 'Måned'/'Måned kort navn'.";
    _Table.AddDataColumn("Måned", "Måned", "4. Måned", DataType.String)
        .Description = "Fullt månedsnavn (f.eks. 'januar').";
    _Table.AddDataColumn("Måned kort navn", "Måned kort navn", "4. Måned", DataType.String)
        .Description = "3-bokstavs forkortet månedsnavn (f.eks. 'jan').";
    _Table.AddDataColumn("År måned", "År måned", "4. Måned", DataType.String)
        .Description = "År og måned som tekst, i formatet 'åååå-MM' (f.eks. '2025-01').";
    _Table.AddDataColumn("År måned no", "År måned no", "4. Måned", DataType.Int64)
        .Description = "År og måned som et sorterbart heltall, i formatet ÅÅÅÅMM (f.eks. 202501), brukt til å sortere 'År måned'/'Måned år'/'År måned filter'.";
    _Table.AddDataColumn("Måned år", "Måned år", "4. Måned", DataType.String)
        .Description = "Forkortet månedsnavn og år (f.eks. 'jan-2025').";

    _Table.AddDataColumn("Kvartal", "Kvartal", "3. Kvartal", DataType.String)
        .Description = "Kvartalet i året, som tekst (f.eks. 'Q1'), brukt til å sortere 'År kvartal'.";
    _Table.AddDataColumn("Kvartal nr", "Kvartal nr", "3. Kvartal", DataType.Int64)
        .Description = "Kvartalnummeret innenfor året (1-4).";
    _Table.AddDataColumn("År kvartal", "År kvartal", "3. Kvartal", DataType.String)
        .Description = "År og kvartal som tekst (f.eks. '2025 Q1').";

    _Table.AddDataColumn("ISO Uke", "ISO Uke", "5. Uke", DataType.Int64)
        .Description = "ISO 8601-ukenummeret innenfor ISO-året (1-53).";
    _Table.AddDataColumn("ISO år", "ISO år", "5. Uke", DataType.Int64)
        .Description = "ISO 8601-året (kan avvike fra kalenderåret for datoer nær årsskiftet).";
    _Table.AddDataColumn("ISO år uke", "ISO år uke", "5. Uke", DataType.String)
        .Description = "ISO 8601-år og -uke som tekst, i formatet 'åååå-Uuu' (f.eks. '2025-U05'), brukt til å sortere 'Uke filter'.";

    _Table.AddDataColumn("Dag", "Dag", "7. Dag", DataType.String)
        .Description = "Ukedagsnavnet (f.eks. 'mandag').";
    _Table.AddDataColumn("Dag i uke", "Dag i uke", "7. Dag", DataType.Int64)
        .Description = "Dagnummeret innenfor uken (1 = mandag .. 7 = søndag), brukt til å sortere 'Dag'.";
    _Table.AddDataColumn("Dag i måned", "Dag i måned", "7. Dag", DataType.Int64)
        .Description = "Dagnummeret innenfor måneden (1-31).";

    _Table.AddDataColumn("År filter", "År filter", "1. Filter", DataType.String)
        .Description = "'Inneværende år' for inneværende år, ellers årstallet som tekst - til bruk i filtre/slicere.";
    _Table.AddDataColumn("År måned filter", "År måned filter", "1. Filter", DataType.String)
        .Description = "'Denne perioden' for inneværende år/måned, ellers 'År måned' - til bruk i filtre/slicere.";
    _Table.AddDataColumn("Uke filter", "Uke filter", "1. Filter", DataType.String)
        .Description = "'Denne uken' for inneværende ISO-år/-uke, ellers 'ISO år uke' - til bruk i filtre/slicere.";
    _Table.AddDataColumn("Dato filter", "Dato filter", "1. Filter", DataType.String)
        .Description = "'Today'/'Yesterday' for disse datoene, ellers datoen som 'yyyy-MM-dd'-tekst - til bruk i filtre/slicere.";

    _Table.AddDataColumn("er_Historie", "er_Historie", "8. Boolean", DataType.Int64)
        .Description = "1 hvis datoen er før i dag, ellers 0.";
    _Table.AddDataColumn("er_Helg", "er_Helg", "8. Boolean", DataType.Int64)
        .Description = "1 hvis datoen faller på en lørdag eller søndag, ellers 0.";

    _key.IsHidden = true;
    _key.IsAvailableInMDX = false;

    Model.Tables[_calendarTableName].DataCategory = "Time";

    // ---- Sort by (slicers, months, day names) ----
    _Table.Columns["Måned"].SortByColumn         = _Table.Columns["Måned nr"];
    _Table.Columns["Måned kort navn"].SortByColumn = _Table.Columns["Måned nr"];
    _Table.Columns["Måned år"].SortByColumn      = _Table.Columns["År måned no"];
    _Table.Columns["År måned"].SortByColumn      = _Table.Columns["År måned no"];
    _Table.Columns["Dag"].SortByColumn           = _Table.Columns["Dag i uke"];

    _Table.Columns["År filter"].SortByColumn       = _Table.Columns["År"];
    _Table.Columns["År måned filter"].SortByColumn = _Table.Columns["År måned no"];
    _Table.Columns["Uke filter"].SortByColumn       = _Table.Columns["ISO år uke"];
    _Table.Columns["Dato filter"].SortByColumn       = _Table.Columns["Dato"];
}
catch (Exception ex)
{
    Info($"Failed to create Date-table.\n\nReason: {ex.Message}");
    return;
}

// Function to create a form with radio buttons
public class SelectionForm : Form
{
    public string SelectedOption { get; private set; } = null;
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
        string _MExpressionStart = $@"
let
    Source = Date.StartOfYear(Date.AddYears(DateTime.LocalNow(), {_yearsFromStart})),
    #""Extracted Date"" = Date.From(Source)
in
    #""Extracted Date""
";
        try
        {
            var sharedExpression = Model.AddExpression(expressionStartName, _MExpressionStart);
            sharedExpression.SetAnnotation("Start date", "Import");
            Info($"Created Shared Expression '{expressionStartName}' successfully.");
        }
        catch (Exception ex)
        {
            Info($"Failed to create Shared Expression '{expressionStartName}'.\n\nReason: {ex.Message}");
        }
    }
    else
    {
        Info($"Skipped creation: Shared Expression '{expressionStartName}' already exists.");
    }

    // Create Relative EndDate
    string expressionEndName = "End date";
    if (!Model.Expressions.Contains(expressionEndName))
    {
        string _MExpressionEnd = $@"
let
    Source = Date.EndOfYear(Date.AddYears(DateTime.LocalNow(), {_yearsFromEnd})),
    #""Extracted Date"" = Date.From(Source)
in
    #""Extracted Date""
";
        try
        {
            var sharedExpression = Model.AddExpression(expressionEndName, _MExpressionEnd);
            sharedExpression.SetAnnotation("End date", "Import");
            Info($"Created Shared Expression '{expressionEndName}' successfully.");
        }
        catch (Exception ex)
        {
            Info($"Failed to create Shared Expression '{expressionEndName}'.\n\nReason: {ex.Message}");
        }
    }
    else
    {
        Info($"Skipped creation: Shared Expression '{expressionEndName}' already exists.");
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
    string _formattedEarliestDate = $"#\"{tableNameEarliest}\"{columnNameEarliest}";

    string expressionEarliestName = "Start date";

    if (!Model.Expressions.Contains(expressionEarliestName))
    {
        string _MExpressionEarliest = $@"
let
    Source = List.Min({_formattedEarliestDate}),
    #""Extracted Date"" = Date.From(Source),
    #""Calculated Start of Year"" = Date.StartOfYear(#""Extracted Date"")
in
    #""Calculated Start of Year""
";
        try
        {
            var sharedExpression = Model.AddExpression(expressionEarliestName, _MExpressionEarliest);
            sharedExpression.SetAnnotation("Start date", "Import");
            Info($"Created Shared Expression '{expressionEarliestName}' successfully.");
        }
        catch (Exception ex)
        {
            Info($"Failed to create Shared Expression '{expressionEarliestName}'.\n\nReason: {ex.Message}");
        }
    }
    else
    {
        Info($"Skipped creation: Shared Expression '{expressionEarliestName}' already exists.");
    }

    // Select Latest Date Column
    string _LatestDate = SelectColumn(_dateColumns, null, "Select the Column with the Latest Date:")
        .DaxObjectFullName;

    // Extract Table Name and Column Name
    var tableNameLatest = _LatestDate.Split('[')[0].Trim('\'');  // Remove single quotes around table name
    var columnNameLatest = _LatestDate.Substring(_LatestDate.IndexOf('[')); // Keep column reference as is

    // Format table name
    string _formattedLatestDate = $"#\"{tableNameLatest}\"{columnNameLatest}";

    string expressionLatestName = "End date";

    if (!Model.Expressions.Contains(expressionLatestName))
    {
        string _MExpressionLatest = $@"
let
    Source = List.Max({_formattedLatestDate}),
    #""Extracted Date"" = Date.From(Source),
    #""Calculated End of Year"" = Date.EndOfYear(#""Extracted Date"")
in
    #""Calculated End of Year""
";
        try
        {
            var sharedExpression = Model.AddExpression(expressionLatestName, _MExpressionLatest);
            sharedExpression.SetAnnotation("End date", "Import");
            Info($"Created Shared Expression '{expressionLatestName}' successfully.");
        }
        catch (Exception ex)
        {
            Info($"Failed to create Shared Expression '{expressionLatestName}'.\n\nReason: {ex.Message}");
        }
    }
    else
    {
        Info($"Skipped creation: Shared Expression '{expressionLatestName}' already exists.");
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
    Info($"Function '{functionName}' already exists. Skipping creation.");
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

    Info($"Created function '{functionName}' successfully.");
}

// ============================================================================
// STEP 2: Mark the table as a Date table
// (merged in directly - no reload, no measures, no relationship, hard-coded to "Dato")
// ============================================================================

try
{
    string _dateColumnName = "Dato";

    if (!_Table.Columns.Contains(_dateColumnName))
    {
        Error($"Column '{_dateColumnName}' not found in table '{_calendarTableName}'. Cannot mark as date table.");
    }
    else
    {
        _Table.DataCategory = "Time";
        _Table.Columns[_dateColumnName].IsKey = true;

        // Table group should be "Dimension"
        _Table.TableGroup = "Dimension";

        // Best practice: prevent numeric columns from being summarized
        foreach (var column in _Table.Columns)
        {
            if (column.DataType == DataType.Int64)
                column.SummarizeBy = AggregateFunction.None;
        }

        // Best practice: apply a consistent short date format string to date columns
        foreach (var column in _Table.Columns)
        {
            if (column.DataType == DataType.DateTime)
                column.FormatString = "dd.MM.yyyy";
        }

        Info($"'{_calendarTableName}' marked as date table using column '{_dateColumnName}'.");
    }
}
catch (Exception ex)
{
    Error($"Error marking '{_calendarTableName}' as date table: {ex.Message}");
}

// ============================================================================
// STEP 3: Add calendar options (Gregoriansk + Ukentlig)
// ============================================================================

try
{
    // Local helper to safely fetch a column by name
    Func<string, Column> col = (name) => _Table.Columns.Contains(name) ? _Table.Columns[name] : null;

    // ----- Gregoriansk calendar (standard calendar) -----
    var calendar1 = _Table.AddCalendar("Gregoriansk");
    calendar1.Description = "Standard kalender ved bruk av kalenderalternativ";

    calendar1.AddTimeUnit(TimeUnit.Year, col("År"), col("År filter"));
    calendar1.AddTimeUnit(TimeUnit.Quarter, col("År kvartal"));
    calendar1.AddTimeUnit(TimeUnit.QuarterOfYear, col("Kvartal nr"), col("Kvartal"));
    calendar1.AddTimeUnit(TimeUnit.Month, col("År måned no"), col("År måned"), col("År måned filter"), col("Måned år"));
    calendar1.AddTimeUnit(TimeUnit.MonthOfYear, col("Måned nr"), col("Måned"), col("Måned kort navn"));
    calendar1.AddTimeUnit(TimeUnit.Date, col("Dato"), col("Dato filter"), col("Dato_Id"));

    Info("Created calendar option 'Gregoriansk'.");

    // ----- Ukentlig calendar (ISO week-based) -----
    var calendar2 = _Table.AddCalendar("Ukentlig");
    calendar2.Description = "Ukebasert (ISO) kalender ved bruk av kalenderalternativ";

    calendar2.AddTimeUnit(TimeUnit.Year, col("ISO år"));
    calendar2.AddTimeUnit(TimeUnit.Week, col("ISO år uke"), col("Uke filter"));
    calendar2.AddTimeUnit(TimeUnit.WeekOfYear, col("ISO Uke"));
    calendar2.AddTimeUnit(TimeUnit.Date, col("Dato"), col("Dato filter"), col("Dato_Id"));

    Info("Created calendar option 'Ukentlig'.");
}
catch (Exception ex)
{
    Error($"Error creating calendar options: {ex.Message}");
}

// ============================================================================
// STEP 4: Relationships to other tables (optional)
// ============================================================================
// Dialog A - pick a mode: skip, connect via the "Dato" (DateTime) column, or connect via the
// "Dato_Id" (Int64) surrogate key column. Either mode reuses the same table->column picker (Dialog B)
// for the ACTIVE relationship (max one column per table) and, optionally, a second pass (Dialog C) to
// add further INACTIVE relationships (multiple columns per table allowed).
string _relationshipMode = "Skip";
{
    var modeForm = new Form { Text = "Relationships to other tables", Width = 480, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    var modeFlow = new FlowLayoutPanel { Padding = new Padding(15), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    modeForm.Controls.Add(modeFlow);

    modeFlow.Controls.Add(new Label { Text = "Create relationships from other tables to this calendar table?", AutoSize = true, Margin = new Padding(0, 0, 0, 10) });
    var rbSkip = new RadioButton { Text = "Skip relationships", AutoSize = true, Checked = true };
    var rbDate = new RadioButton { Text = $"Connect with Date-format (uses '{_calendarTableName}'[Dato])", AutoSize = true };
    var rbInt  = new RadioButton { Text = $"Connect with Int-format (uses '{_calendarTableName}'[Dato_Id])", AutoSize = true };
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
                            if (sibling.Tag is Column siblingCol) checkedNames.Remove(siblingCol.DaxObjectFullName);
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

    string oneSideColumnName = _relationshipMode == "Date" ? "Dato" : "Dato_Id";
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
        Info($"No {typeLabel} columns found on other tables - skipping relationship creation.");
    }
    else
    {
        Func<Column, Column, bool> RelationshipExists = (fromCol, toCol) =>
            Model.Relationships.Any(r => r.FromColumn == fromCol && r.ToColumn == toCol);

        var activeSelections = PickColumnsDialog(candidateColumns, true,
            "Select the ACTIVE relationship column",
            $"For each table below, optionally check ONE {typeLabel} column to create an ACTIVE relationship to '{_calendarTableName}'[{oneSideColumnName}]. Leave a table unchecked to skip it.");

        foreach (var fromCol in activeSelections)
        {
            if (RelationshipExists(fromCol, oneSideColumn))
            {
                Info($"Relationship from '{fromCol.DaxObjectFullName}' to '{oneSideColumn.DaxObjectFullName}' already exists - skipping.");
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
                    $"Select any number of additional {typeLabel} columns (multiple per table allowed) to create as INACTIVE relationships to '{_calendarTableName}'[{oneSideColumnName}].");

                foreach (var fromCol in inactiveSelections)
                {
                    if (RelationshipExists(fromCol, oneSideColumn))
                    {
                        Info($"Relationship from '{fromCol.DaxObjectFullName}' to '{oneSideColumn.DaxObjectFullName}' already exists - skipping.");
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

// ============================================================================
// STEP 5: Add English ("en-US") translations for the table and its columns
// ============================================================================
try
{
    string translationCulture = "en-US";

    if (!Model.Cultures.Contains(translationCulture))
    {
        Culture.CreateNew(Model, translationCulture);
        Info($"Culture '{translationCulture}' did not exist and was created.");
    }

    _Table.TranslatedNames[translationCulture] = "Calendar";

    // Norwegian column name -> (English name, English description), taken from
    // "1 - Create English Calendar (combined).cs".
    var englishTranslations = new Dictionary<string, (string Name, string Description)>
    {
        { "Dato",             ("Date", "The calendar date.") },
        { "Dato_Id",          ("Date_Key", "Integer surrogate key for the date, in YYYYMMDD format (e.g. 20250131).") },
        { "År",               ("Year", "The calendar year number (e.g. 2025).") },
        { "Måned nr",         ("Month no", "The month number within the year (1-12), used to sort 'Month name'/'Month short name'.") },
        { "Måned",            ("Month name", "The full month name (e.g. 'January').") },
        { "Måned kort navn",  ("Month short name", "The 3-letter abbreviated month name (e.g. 'Jan').") },
        { "År måned",         ("Year month", "The year and month as text, in 'yyyy-MM' format (e.g. '2025-01').") },
        { "År måned no",      ("Year month no", "The year and month as a sortable integer, in YYYYMM format (e.g. 202501), used to sort 'Year month'/'Month name Year'/'Year month slicer'.") },
        { "Måned år",         ("Month name Year", "The abbreviated month name and year (e.g. 'Jan-2025').") },
        { "Kvartal",          ("Quarter", "The quarter of the year, as text (e.g. 'Q1'), used to sort 'Year quarter'.") },
        { "Kvartal nr",       ("Quarter no", "The quarter number within the year (1-4).") },
        { "År kvartal",       ("Year quarter", "The year and quarter as text (e.g. '2025 Q1').") },
        { "ISO Uke",          ("ISO week", "The ISO 8601 week number within the ISO year (1-53).") },
        { "ISO år",           ("ISO year", "The ISO 8601 year (can differ from the calendar year for dates near year-end/-start).") },
        { "ISO år uke",       ("ISO year week", "The ISO 8601 year and week as text, in 'yyyy-Uww' format (e.g. '2025-U05'), used to sort 'Week slicer'.") },
        { "Dag",              ("Day name", "The day of week name (e.g. 'Monday').") },
        { "Dag i uke",        ("Day no of Week", "The day number within the week (1 = Monday .. 7 = Sunday), used to sort 'Day name'.") },
        { "Dag i måned",      ("Day no of Month", "The day number within the month (1-31).") },
        { "År filter",        ("Year slicer", "'Inneværende år' for the current year, otherwise the year number as text - for use in slicers.") },
        { "År måned filter",  ("Year month slicer", "'Denne perioden' for the current year/month, otherwise 'Year month' - for use in slicers.") },
        { "Uke filter",       ("Week slicer", "'Denne uken' for the current ISO year/week, otherwise 'ISO year week' - for use in slicers.") },
        { "Dato filter",      ("Date slicer", "'Today'/'Yesterday' for those dates, otherwise the date as 'yyyy-MM-dd' text - for use in slicers.") },
        { "er_Historie",      ("is_History", "1 if the date is before today, otherwise 0.") },
        { "er_Helg",          ("is_Weekend", "1 if the date falls on a Saturday or Sunday, otherwise 0.") },
    };

    foreach (var kvp in englishTranslations)
    {
        if (_Table.Columns.Contains(kvp.Key))
        {
            var column = _Table.Columns[kvp.Key];
            column.TranslatedNames[translationCulture] = kvp.Value.Name;
            column.TranslatedDescriptions[translationCulture] = kvp.Value.Description;
        }
    }

    Info($"Added '{translationCulture}' translations for table '{_calendarTableName}' and its columns.");
}
catch (Exception ex)
{
    Error($"Error adding translations: {ex.Message}");
}

Info("Calendar table creation, date-table setup, calendar options, relationships and translations completed successfully!");
