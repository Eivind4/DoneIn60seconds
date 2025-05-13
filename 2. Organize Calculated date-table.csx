// Script to perform maintenance on the date-table in the Contoso dataset.
// Needs a table on the same format with the same columns, or it needs to be modified


// Hide the 'Running Macro' spinbox
ScriptHelper.WaitFormVisible = false;

if (Selected.Table == null)
{
    Error("No table selected in the UI. Please select a date table before running the script.");
    return;
}

string _CalendarTable = Selected.Table.Name;

try
{
    var table = Model.Tables[_CalendarTable];

    // Sorting
     table.Columns["Month name"].SortByColumn = table.Columns["Month no"];
        table.Columns["Month short name"].SortByColumn = table.Columns["Month no"];
        table.Columns["Year month"].SortByColumn = table.Columns["Year month no"];
        table.Columns["Month name Year"].SortByColumn = table.Columns["Year month no"];
        table.Columns["Day name"].SortByColumn = table.Columns["Day no of week"];
        table.Columns["Year slicer"].SortByColumn = table.Columns["Year"];
        table.Columns["Year month slicer"].SortByColumn = table.Columns["Year month no"];
        table.Columns["Week slicer"].SortByColumn = table.Columns["ISO year week"];
        table.Columns["Date slicer"].SortByColumn = table.Columns["Date"];
    
        // Sisplay Folders
    string slicerFolder = "1. Slicer";
    string yearFolder = "2. Year";
    string quarterFolder = "3. Quarter";
    string monthFolder = "4. Month";
    string weekFolder = "5. Week";
    string dateFolder = "6. Date";
    string dayFolder = "7. Day";
    string booleanFolder = "8. Boolean";

    // Slicer
    table.Columns["Year slicer"].DisplayFolder = slicerFolder;
    table.Columns["Year month slicer"].DisplayFolder = slicerFolder;
    table.Columns["Week slicer"].DisplayFolder = slicerFolder;
    table.Columns["Date slicer"].DisplayFolder = slicerFolder;

    // Year
    table.Columns["Year"].DisplayFolder = yearFolder;

    // Quarter
    table.Columns["Quarter"].DisplayFolder = quarterFolder;
    table.Columns["Quarter no"].DisplayFolder = quarterFolder;
    table.Columns["Year quarter"].DisplayFolder = quarterFolder;

    // Month
    table.Columns["Month no"].DisplayFolder = monthFolder;
    table.Columns["Month name"].DisplayFolder = monthFolder;
    table.Columns["Month short name"].DisplayFolder = monthFolder;
    table.Columns["Year month"].DisplayFolder = monthFolder;
    table.Columns["Year month no"].DisplayFolder = monthFolder;
    table.Columns["Month name Year"].DisplayFolder = monthFolder;

    // Week
    table.Columns["ISO week"].DisplayFolder = weekFolder;
    table.Columns["ISO year"].DisplayFolder = weekFolder;
    table.Columns["ISO year week"].DisplayFolder = weekFolder;

    // Date
    table.Columns["Date"].DisplayFolder = dateFolder;
    table.Columns["Date_Key"].DisplayFolder = dateFolder;

    // Day
    table.Columns["Day name"].DisplayFolder = dayFolder;
    table.Columns["Day no of Week"].DisplayFolder = dayFolder;
    table.Columns["Day no of Month"].DisplayFolder = dayFolder;

    // Boolean
    table.Columns["is_History"].DisplayFolder = booleanFolder;
    table.Columns["is_Weekend"].DisplayFolder = booleanFolder;

    // Format Strings
    table.Columns["Date"].FormatString = "mm/dd/yyyy";
    table.Columns["Date_Key"].FormatString = "0";
    table.Columns["Day no of Month"].FormatString = "0";
    table.Columns["Day no of Week"].FormatString = "0";
    table.Columns["is_History"].FormatString = "0";
    table.Columns["ISO week"].FormatString = "0";
    table.Columns["ISO year"].FormatString = "0";
    table.Columns["Month no"].FormatString = "0";
    table.Columns["Quarter no"].FormatString = "0";
    table.Columns["Year"].FormatString = "0";
    table.Columns["Year month no"].FormatString = "0";

    // Hide Unused Columns
    table.Columns["Date_Key"].IsHidden = true;
    table.Columns["Date_Key"].IsAvailableInMDX = false;

    // Column Descriptions
    table.Columns["Date"].Description = "YYYY-MM-DD";
    table.Columns["Date_Key"].Description = "YYYYMMDD";
    table.Columns["Date slicer"].Description = "Today, yesterday or date";

    table.Columns["Year"].Description = "YYYY";
    table.Columns["Year slicer"].Description = "Current year, or YYYY";

    table.Columns["Quarter no"].Description = "1, 2, 3 or 4";
    table.Columns["Quarter"].Description = "e.g., Q1";
    table.Columns["Year quarter"].Description = "2024 Q1";

    table.Columns["Month no"].Description = "MM";
    table.Columns["Month name"].Description = "e.g., January";
    table.Columns["Month short name"].Description = "e.g., Jan";
    table.Columns["Year month"].Description = "YYYY-MM";
    table.Columns["Year month no"].Description = "YYYYMM";
    table.Columns["Month name Year"].Description = "e.g., Jan 2024";
    table.Columns["Year month slicer"].Description = "Current month or YYYY-MM";

    table.Columns["ISO year"].Description = "ISO year (YYYY)";
    table.Columns["ISO week"].Description = "ISO week number";
    table.Columns["ISO year week"].Description = "e.g., 2024 W01";
    table.Columns["Week slicer"].Description = "Current week or ISO year week";

    table.Columns["Day no of Month"].Description = "Day number 1-31";
    table.Columns["Day no of Week"].Description = "0=Monday to 6=Sunday";
    table.Columns["Day name"].Description = "e.g., Monday";

    table.Columns["is_Weekend"].Description = "1 if weekend, else 0";
    table.Columns["is_History"].Description = "1 if today or earlier";

    // 8. Table description
    table.Description = "Date-table";
}
catch
{
    Error("No selected date-table or see if column definitions are matching => No formatting");
}
