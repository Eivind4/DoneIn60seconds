// Adapted from https://github.com/data-goblin/powerbi-macguyver-toolbox/blob/main/tabular-editor-scripts/csharp-scripts/import-model-only/create-date-table.csx
// This script works in Tabular Editor 2

// It is modified by including other naming convention and other columns, like ISO Week

// 1. Select the column with the earliest date
// 2. Select the column with the latest date
// 3. Script generates a calculated table with rich date metadata

var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime).ToList();

string _EarliestDateInput;
try
{
    _EarliestDateInput = SelectColumn(_dateColumns, null, "Select the Column with the Earliest Date:").DaxObjectFullName;
}
catch
{
    Error("Earliest column not selected! Script stopped.");
    return;
}

string _LatestDateInput;
try
{
    _LatestDateInput = SelectColumn(_dateColumns, null, "Select the Column with the Latest Date:").DaxObjectFullName;
}
catch
{
    Error("Latest column not selected! Script stopped.");
    return;
}

try
{
    string _RefDateMeasureLogic = string.Format("CALCULATE ( MAX ( {0} ), REMOVEFILTERS ( ) )", _LatestDateInput);

    string _DateDaxExpression = string.Format(@"
VAR _Today = TODAY ()

VAR _EarliestDate = DATE ( YEAR ( MIN ( {0} ) ), 1, 1 )
VAR _EarliestDate_Safe = MIN ( _EarliestDate, DATE ( YEAR ( _Today ), 1, 1 ) )

VAR _LatestDate = DATE ( YEAR ( MAX ( {1} ) ), 12, 31 )
VAR _LatestDate_Safe = MAX ( _LatestDate, DATE ( YEAR ( _Today ), 12, 31 ) )

VAR _Base_Calendar = CALENDAR ( _EarliestDate_Safe, _LatestDate_Safe )

VAR _IntermediateResult =
    ADDCOLUMNS (
        _Base_Calendar,
        ""Year"", YEAR ( [Date] ),
        ""Quarter no"", QUARTER ( [Date] ),
        ""Quarter"", ""Q"" & CONVERT ( QUARTER ( [Date] ), STRING ),
        ""Year quarter"", CONVERT ( YEAR ( [Date] ), STRING ) & "" Q"" & CONVERT ( QUARTER ( [Date] ), STRING ),
        ""Month no"", MONTH ( [Date] ),
        ""Month name"", FORMAT ( [Date], ""MMMM"" ),
        ""Month short name"", FORMAT ( [Date], ""MMM"" ),
        ""Year month"", CONVERT ( YEAR ( [Date] ), STRING ) & ""-"" & FORMAT ( MONTH ( [Date] ), ""00"" ),
        ""Year month no"", YEAR ( [Date] ) * 100 + MONTH ( [Date] ),
        ""Month name Year"", FORMAT ( [Date], ""MMM YY"" ),
        ""ISO week"", WEEKNUM ( [Date], 21 ),
        ""ISO year"", YEAR ( [Date] + (26 - WEEKNUM ( [Date], 21 )) ),
        ""Day name"", FORMAT ( [Date], ""DDDD"" ),
        ""Day no of Week"", WEEKDAY ( [Date], 3 ),
        ""Day no of Month"", DAY ( [Date] ),
        ""Date_Key"", YEAR ( [Date] ) * 10000 + MONTH ( [Date] ) * 100 + DAY ( [Date] ),
        ""Year slicer"", IF ( YEAR ( [Date] ) = YEAR ( TODAY() ), ""Current Year"", CONVERT ( YEAR ( [Date] ), STRING ) ),
        ""Year month slicer"", IF ( AND ( MONTH ( [Date] ) = MONTH ( TODAY() ), YEAR ( [Date] ) = YEAR ( TODAY() ) ), ""Current month"", CONVERT ( YEAR ( [Date] ), STRING ) & ""-"" & FORMAT ( MONTH ( [Date] ), ""00"" ) ),
        ""Date slicer"", IF ( [Date] = TODAY(), ""Today"", IF ( [Date] = TODAY() - 1, ""Yesterday"", CONVERT ( [Date], STRING ) ) ),
        ""is_History"", IF ( [Date] <= TODAY(), 1, 0 ),
        ""is_Weekend"", WEEKDAY ( [Date], 3 ) IN {{5,6}}
    )

VAR _Result =
    ADDCOLUMNS (
        _IntermediateResult,
        ""ISO year week"", CONVERT ( [ISO year], STRING ) & ""-W"" & FORMAT ( [ISO week], ""00"" ),
        ""Week slicer"",
            VAR CurrentDate = TODAY()
            VAR ThursdayOfWeek = CurrentDate + (3 - WEEKDAY ( CurrentDate, 2 ))
            VAR ThursdayToday = YEAR ( ThursdayOfWeek )
            RETURN IF (
                AND (
                    ThursdayToday = [ISO year],
                    WEEKNUM ( TODAY(), 21 ) = [ISO week]
                ),
                ""Current week"",
                CONVERT ( [ISO year], STRING ) & ""-W"" & FORMAT ( [ISO week], ""00"" )
            )
    )
RETURN _Result", _EarliestDateInput, _LatestDateInput);

    // Create calculated table
    var _date = Model.AddCalculatedTable("Calendar", _DateDaxExpression);

    // Mark as date table and assign group
    _date.DataCategory = "Time";

    // Remove summarization from int columns
    foreach (var column in _date.Columns.Where(c => c.DataType == DataType.Int64))
        column.SummarizeBy = AggregateFunction.None;

    // Apply date format to datetime columns
    foreach (var column in _date.Columns.Where(c => c.DataType == DataType.DateTime))
        column.FormatString = "mm/dd/yyyy";

    Info(string.Format("Created a new, organized 'Calendar' table.\nEarliest Date: {0}\nLatest Date: {1}", _EarliestDateInput, _LatestDateInput));
}
catch (Exception ex)
{
    Error("An unexpected error occurred while creating the Calendar table: " + ex.Message);
}
