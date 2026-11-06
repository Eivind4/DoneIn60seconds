// The purpose of this macro is to run it as part of the create calendar macro, if write operations are allowed. 
// If you know you want to write a table using English standard format, this can be run standalone.

/*
 * Title: Create English Calendar
 * 
 * Author: Eivind Haugen
 * 
 * This script, when executed, will write a new table (currently for XLMA-endpoint (Premium) or PBIP. 
 * 
 * It will let you select whether you want to select the start and from date based on either a relative date to current date or 
 * based on a selected start- and end-date from your fact tables. Shared expressions will then be created. It also creates shared expressions used for formatting
 * as the basis is on what can be used as format without space in the source.
 */

#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System;
using System.Drawing;
using System.Windows.Forms;

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;


// Creating the English Calendar

string _calendarTableName = Interaction.InputBox("Provide the name of the calendar table table", "Table Name", "Calendar");

// Check if the "Calendar" table already exists
if (Model.Tables.Contains(_calendarTableName))
{
    Info($"Table '{_calendarTableName}' already exists. Skipping creation.");
    return;
}

// Create the "Calendar" table
var _Table = Model.AddTable(_calendarTableName, false);

// Define the M expression for partition (Fixing Syntax Issues)
    string _MPartition = @"
let
    Source = List.Dates(#""Start date"", Number.From(#""End date"" - #""Start date"") + 1, #duration(1, 0, 0, 0)),
    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #""Renamed Columns"" = Table.RenameColumns(#""Converted to Table"", {{""Column1"", ""Dato""}}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Renamed Columns"", {{""Dato"", type date}}),

    // Adding calculated columns
    AddedColumns = Table.AddColumn(#""Changed Type"", ""Dato_Id"", each Date.Year([Dato]) * 10000 + Date.Month([Dato]) * 100 + Date.Day([Dato]), Int64.Type),
    
    AddedDateSlicer = Table.AddColumn(AddedColumns, ""Dato_filter"", each 
        if [Dato] = Date.From(DateTime.LocalNow()) then ""Today"" 
        else if [Dato] = Date.AddDays(Date.From(DateTime.LocalNow()), -1) then ""Yesterday"" 
        else Date.ToText([Dato], ""yyyy-MM-dd""), type text),
    
    AddedYear = Table.AddColumn(AddedDateSlicer, ""Aar"", each Date.Year([Dato]), Int64.Type),
    
    AddedYearSlicer = Table.AddColumn(AddedYear, ""Aar_filter"", each 
        if Date.Year([Dato]) = Date.Year(DateTime.LocalNow()) then ""Current Year"" 
        else Text.From(Date.Year([Dato])), type text),
    
    AddedQuarterNo = Table.AddColumn(AddedYearSlicer, ""Kvartal_nr"", each Date.QuarterOfYear([Dato]), Int64.Type),
    
    AddedQuarter = Table.AddColumn(AddedQuarterNo, ""Kvartal"", each ""Q"" & Text.From(Date.QuarterOfYear([Dato])), type text),
    
    AddedYearQuarter = Table.AddColumn(AddedQuarter, ""Aar_kvartal"", each 
        Text.From(Date.Year([Dato])) & "" Q"" & Text.From(Date.QuarterOfYear([Dato])), type text),
    
    AddedMonthNo = Table.AddColumn(AddedYearQuarter, ""Maaned_nr"", each Date.Month([Dato]), Int64.Type),
    
    AddedMonthName = Table.AddColumn(AddedMonthNo, ""Maaned"", each Text.Proper(Date.ToText([Dato], ""MMMM"", ""nb-NO"")), type text),
    
    AddedMonthShortName = Table.AddColumn(AddedMonthName, ""Maaned_kort_navn"", each Text.Proper(Text.Start(Date.ToText([Dato], ""MMMM"", ""nb-NO""), 3)), type text),
    
    AddedYearMonth = Table.AddColumn(AddedMonthShortName, ""Aar_maaned"", each Date.ToText([Dato], ""yyyy-MM""), type text),
    AddedYearMonthNo = Table.AddColumn(AddedYearMonth, ""Aar_maaned_no"", each Date.Year([Dato]) * 100 + Date.Month([Dato]), Int64.Type),
    AddedMonthNameYear = Table.AddColumn(AddedYearMonthNo, ""Maaned_aar"", each Text.Combine({Text.Proper(Text.Start(Date.ToText([Dato], ""MMMM"", ""nb-NO""), 3)), Text.From(Date.Year([Dato]), ""en-US"")}, ""-""), type text),
    AddedYearMonthSlicer = Table.AddColumn(AddedMonthNameYear, ""Aar_maaned_filter"", each if Date.Year([Dato]) = Date.Year(DateTime.LocalNow()) and Date.Month([Dato]) = Date.Month(DateTime.LocalNow()) then ""Current month"" else [Aar_maaned], type text),
 ////   AddedMonthStartDate = Table.AddColumn(AddedYearMonthSlicer, ""Maaned_start_dato"", each Date.StartOfMonth([Dato]), type date),
    
 ////   AddedMonthEndDate = Table.AddColumn(AddedMonthStartDate, ""Maaned_slutt_dato"", each Date.EndOfMonth([Dato]), type date),
    #""Inserted Day"" = Table.AddColumn(AddedYearMonthSlicer, ""Dag_i_maaned"", each Date.Day([Dato]), Int64.Type),
    #""Inserted Day of Week"" = Table.AddColumn(#""Inserted Day"", ""Dag_i_uke"", each Date.DayOfWeek([Dato],Day.Monday), Int64.Type),
    #""Inserted Day Name"" = Table.AddColumn(#""Inserted Day of Week"", ""Dag"", each Text.Proper(Date.ToText([Dato], ""dddd"", ""nb-NO"")), type text),
    #""Inserted is_Weekend"" = Table.AddColumn(#""Inserted Day Name"", ""er_Helg"", each if Date.DayOfWeek([Dato]) >= 5 then 1 else 0, Int64.Type),
 ////   #""Inserted is_Last day of month"" = Table.AddColumn(#""Inserted is_Weekend"", ""er_Siste_dag_i_maaned"", each if [Dato] = Date.EndOfMonth([Dato]) then 1 else 0, Int64.Type),
    #""Inserted is History"" = Table.AddColumn(#""Inserted is_Weekend"", ""er_Historie"", each if [Dato] < Date.From(DateTime.LocalNow()) then 1 else 0, Int64.Type),
//    AddedFiscalYear = Table.AddColumn(#""Inserted is History"", ""Fiscal_year"", each 
//        if Date.Month([Dato]) >= 7 then Date.Year([Dato]) + 1 else Date.Year([Dato]), Int64.Type),
    
 //   AddedFiscalMonthNo = Table.AddColumn(AddedFiscalYear, ""Fiscal_month_no"", each 
 //       if Date.Month([Dato]) >= 7 then Date.Month([Dato]) - 6 else Date.Month([Dato]) + 6, Int64.Type),
    
 //   AddedFiscalYearMonth = Table.AddColumn(AddedFiscalMonthNo, ""Fiscal_year_month"", each 
 //       Text.From(if Date.Month([Dato]) >= 7 then Date.Year([Dato]) + 1 else Date.Year([Dato])) & ""-"" & Text.PadStart(Text.From(if Date.Month([Dato]) >= 7 then Date.Month([Dato]) - 6 else Date.Month([Dato]) + 6), 2, ""0""), type text),
    
 //   AddedTertiaryNo = Table.AddColumn(AddedFiscalYearMonth, ""Tertiary_no"", each Number.RoundUp(Date.Month([Dato]) / 4), Int64.Type),
    
 //   AddedYearTertiary = Table.AddColumn(AddedTertiaryNo, ""Year_tertiary"", each 
//        Text.From(Date.Year([Dato])) & ""-0"" & Text.From(Number.RoundUp(Date.Month([Dato]) / 4)), type text),
    
//    AddedHalfyearNo = Table.AddColumn(AddedYearTertiary, ""Halfyear_no"", each Number.RoundUp(Date.Month([Dato]) / 6), Int64.Type),
    
//    AddedYearHalfyear = Table.AddColumn(AddedHalfyearNo, ""Year_halfyear"", each 
 //       Text.From(Date.Year([Dato])) & ""-0"" & Text.From(Number.RoundUp(Date.Month([Dato]) / 6)), type text),
    #""Added ISO_week"" = Table.AddColumn(#""Inserted is History"", ""ISO_Uke"", each let
        CurrentThursday = Date.AddDays([Dato], 3 - Date.DayOfWeek([Dato], Day.Monday)),
        YearCurrThursday = Date.Year(CurrentThursday),
        FirstThursdayOfYear = Date.AddDays(#date(YearCurrThursday, 1, 7), -Date.DayOfWeek(#date(YearCurrThursday, 1, 1), Day.Friday)),
        ISO_Week = Duration.Days(CurrentThursday - FirstThursdayOfYear) / 7 + 1
    in
        ISO_Week, Int64.Type),
    #""Added ISO_year"" = Table.AddColumn(#""Added ISO_week"", ""ISO_Aar"", each Date.Year(Date.AddDays([Dato], 26 - [ISO_Uke])), Int64.Type),
    #""Added ISO year week"" = Table.AddColumn(#""Added ISO_year"", ""ISO_Aar_uke"", each let
        ISO_Year = [ISO_Aar],
        ISO_Week = [ISO_Uke],
        ISO_WeekFormatted = Text.PadStart(Text.From(ISO_Week), 2, ""0""),
        ISO_Year_Week = Text.Combine({"""", Text.From(ISO_Year), ""-W"", ISO_WeekFormatted})
    in
        ISO_Year_Week, type text),
    AddWeekSlicer = Table.AddColumn(#""Added ISO year week"", ""Uke_filter"", each 
        if [ISO_Aar] = Date.Year(DateTime.LocalNow()) and [ISO_Uke] = Date.WeekOfYear(DateTime.LocalNow(), Day.Monday) 
        then ""Current_Week"" 
        else Text.From([ISO_Aar_uke]), 
        type text),
    Calendar = Table.TransformColumnNames(AddWeekSlicer, each RenameColumn(_)),
    FixingAA = Table.TransformColumnNames(Calendar, each Replace_aa(_))
in
    FixingAA";


try
{
    // Add partition
    var partition = _Table.AddMPartition(_calendarTableName, _MPartition);
    partition.Mode = ModeType.Import; // Set refresh policy to Import

    // Adding data columns
 _Table.AddDataColumn("Dato", "Dato", "6. Dato", DataType.DateTime);
var _key = _Table.AddDataColumn("Dato_Id", "Dato_Id", "6. Dato", DataType.String);

_Table.AddDataColumn("År", "År", "2. År", DataType.Int64);

_Table.AddDataColumn("Måned nr", "Måned nr", "4. Måned", DataType.Int64);
_Table.AddDataColumn("Måned", "Måned", "4. Måned", DataType.String);
_Table.AddDataColumn("Måned kort navn", "Måned kort navn", "4. Måned", DataType.String);
_Table.AddDataColumn("År måned", "År måned", "4. Måned", DataType.String);
_Table.AddDataColumn("År måned no", "År måned no", "4. Måned", DataType.Int64);
_Table.AddDataColumn("Måned år", "Måned år", "4. Måned", DataType.String);

_Table.AddDataColumn("Kvartal", "Kvartal", "3. Kvartal", DataType.String);
_Table.AddDataColumn("Kvartal nr", "Kvartal nr", "3. Kvartal", DataType.Int64);
_Table.AddDataColumn("År kvartal", "År kvartal", "3. Kvartal", DataType.String);

_Table.AddDataColumn("ISO Uke", "ISO Uke", "5. Uke", DataType.Int64);
_Table.AddDataColumn("ISO år", "ISO år", "5. Uke", DataType.Int64);
_Table.AddDataColumn("ISO år uke", "ISO år uke", "5. Uke", DataType.String);

_Table.AddDataColumn("Dag", "Dag", "7. Dag", DataType.String);
_Table.AddDataColumn("Dag i uke", "Dag i uke", "7. Dag", DataType.Int64);
_Table.AddDataColumn("Dag i måned", "Dag i måned", "7. Dag", DataType.Int64);

 _Table.AddDataColumn("År filter", "År filter", "1. Filter", DataType.String);
_Table.AddDataColumn("År måned filter", "År måned filter", "1. Filter", DataType.String);
_Table.AddDataColumn("Uke filter", "Uke filter", "1. Filter", DataType.String);
_Table.AddDataColumn("Dato filter", "Dato filter", "1. Filter", DataType.String);

_Table.AddDataColumn("er_Historie", "er_Historie", "8. Boolean", DataType.Int64);
_Table.AddDataColumn("er_Helg", "er_Helg", "8. Boolean", DataType.Int64);

_key.IsHidden = true;
_key.IsAvailableInMDX = false;

    Model.Tables[_calendarTableName].DataCategory = "Time";

 //   Info("Created Date-table successfully.");
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

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

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

// Format table name properly
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

// Select Earliest Date Column
string _LatestDate = SelectColumn(_dateColumns, null, "Select the Column with the Earliest Date:")
    .DaxObjectFullName;

// Extract Table Name and Column Name
var tableNameLatest = _LatestDate.Split('[')[0].Trim('\'');  // Remove single quotes around table name
var columnNameLatest = _LatestDate.Substring(_LatestDate.IndexOf('[')); // Keep column reference as is

// Format table name properly
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

