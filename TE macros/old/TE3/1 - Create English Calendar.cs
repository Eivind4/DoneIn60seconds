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
    //AddedMonthStartDate = Table.AddColumn(AddedYearMonthSlicer, ""Month_start_date"", each Date.StartOfMonth([Date]), type date),
    
    //AddedMonthEndDate = Table.AddColumn(AddedMonthStartDate, ""Month_end_date"", each Date.EndOfMonth([Date]), type date),
    // #""Inserted is_Last day of month"" = Table.AddColumn(#""Inserted is_Weekend"", ""is_Last_day_of_month"", each if [Date] = Date.EndOfMonth([Date]) then 1 else 0, Int64.Type),
    //#""Inserted Day"" = Table.AddColumn(AddedMonthEndDate, ""Day_no_of_month"", each Date.Day([Date]), Int64.Type),
    #""Inserted Day"" = Table.AddColumn(AddedYearMonthSlicer, ""Day_no_of_Month"", each Date.Day([Date]), Int64.Type),
    #""Inserted Day of Week"" = Table.AddColumn(#""Inserted Day"", ""Day_no_of_Week"", each Date.DayOfWeek([Date],Day.Monday), Int64.Type),
    #""Inserted Day Name"" = Table.AddColumn(#""Inserted Day of Week"", ""Day_name"", each Date.DayOfWeekName([Date]), type text),
    #""Inserted is_Weekend"" = Table.AddColumn(#""Inserted Day Name"", ""is_Weekend"", each if Date.DayOfWeek([Date]) >= 5 then 1 else 0, Int64.Type),
   
    #""Inserted is History"" = Table.AddColumn(#""Inserted is_Weekend"", ""is_History"", each if [Date] < Date.From(DateTime.LocalNow()) then 1 else 0, Int64.Type),
   // AddedFiscalYear = Table.AddColumn(#""Inserted is History"", ""Fiscal_year"", each 
  //      if Date.Month([Date]) >= 7 then Date.Year([Date]) + 1 else Date.Year([Date]), Int64.Type),
    
  //  AddedFiscalMonthNo = Table.AddColumn(AddedFiscalYear, ""Fiscal_month_no"", each 
  //      if Date.Month([Date]) >= 7 then Date.Month([Date]) - 6 else Date.Month([Date]) + 6, Int64.Type),
    
  //  AddedFiscalYearMonth = Table.AddColumn(AddedFiscalMonthNo, ""Fiscal_year_month"", each 
 //       Text.From(if Date.Month([Date]) >= 7 then Date.Year([Date]) + 1 else Date.Year([Date])) & ""-"" & Text.PadStart(Text.From(if Date.Month([Date]) >= 7 then Date.Month([Date]) - 6 else Date.Month([Date]) + 6), 2, ""0""), type text),
    
  //  AddedTertiaryNo = Table.AddColumn(AddedFiscalYearMonth, ""Tertiary_no"", each Number.RoundUp(Date.Month([Date]) / 4),Int64.Type),
    
 //   AddedYearTertiary = Table.AddColumn(AddedTertiaryNo, ""Year_tertiary"", each 
  //      Text.From(Date.Year([Date])) & ""-0"" & Text.From(Number.RoundUp(Date.Month([Date]) / 4)), type text),
    
 //   AddedHalfyearNo = Table.AddColumn(AddedYearTertiary, ""Halfyear_no"", each Number.RoundUp(Date.Month([Date]) / 6), Int64.Type),
    
 //   AddedYearHalfyear = Table.AddColumn(AddedHalfyearNo, ""Year_halfyear"", each 
 //       Text.From(Date.Year([Date])) & ""-0"" & Text.From(Number.RoundUp(Date.Month([Date]) / 6)), type text),
 //   #""Added ISO_week"" = Table.AddColumn(AddedYearHalfyear, ""ISO_week"", each let
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
    _Table.AddDataColumn("Date", "Date", "6. Date", DataType.DateTime);
    var _key = _Table.AddDataColumn("Date_Key", "Date_Key", "6. Date", DataType.Int64);

    _Table.AddDataColumn("Year", "Year", "2. Year", DataType.Int64);

    _Table.AddDataColumn("Month no", "Month no", "4. Month", DataType.Int64);
    _Table.AddDataColumn("Month name", "Month name", "4. Month", DataType.String);
    _Table.AddDataColumn("Month short name", "Month short name", "4. Month", DataType.String);
    _Table.AddDataColumn("Year month", "Year month", "4. Month", DataType.String);
    _Table.AddDataColumn("Year month no", "Year month no", "4. Month", DataType.Int64);
    _Table.AddDataColumn("Month name Year", "Month name Year", "4. Month", DataType.String);

    _Table.AddDataColumn("Quarter", "Quarter", "3. Quarter", DataType.String);
    _Table.AddDataColumn("Quarter no", "Quarter no", "3. Quarter", DataType.Int64);
    _Table.AddDataColumn("Year quarter", "Year quarter", "3. Quarter", DataType.String);

    _Table.AddDataColumn("ISO week", "ISO week", "5. Week", DataType.Int64);
    _Table.AddDataColumn("ISO year", "ISO year", "5. Week", DataType.Int64);
    _Table.AddDataColumn("ISO year week", "ISO year week", "5. Week", DataType.String);

    _Table.AddDataColumn("Day name", "Day name", "7. Day", DataType.String);
    _Table.AddDataColumn("Day no of Week", "Day no of Week", "7. Day", DataType.Int64);
    _Table.AddDataColumn("Day no of Month", "Day no of Month", "7. Day", DataType.Int64);

    _Table.AddDataColumn("Year Slicer", "Year Slicer", "1. Slicer", DataType.String);
    _Table.AddDataColumn("Year month slicer", "Year month slicer", "1. Slicer", DataType.String);
    _Table.AddDataColumn("Week slicer", "Week slicer", "1. Slicer", DataType.String);
    _Table.AddDataColumn("Date slicer", "Date slicer", "1. Slicer", DataType.String);

    _Table.AddDataColumn("is_History", "is_History", "8. Boolean", DataType.Int64);
    _Table.AddDataColumn("is_Weekend", "is_Weekend", "8. Boolean", DataType.Int64);

    _key.IsHidden = true;
    _key.IsAvailableInMDX = false;

    Model.Tables[_calendarTableName].DataCategory = "Time";

  //  Info("Created Date-table successfully.");
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

// Select Earliest Date Column
string _LatestDate = SelectColumn(_dateColumns, null, "Select the Column with the Earliest Date:")
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

