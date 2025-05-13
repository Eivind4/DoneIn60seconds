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
    table.Columns["Year Quarter"].SortByColumn = table.Columns["Year Quarter Number"];
    table.Columns["Year Month Short"].SortByColumn = table.Columns["Year Month Number"];
    table.Columns["Year Month"].SortByColumn = table.Columns["Year Month Number"];
    table.Columns["Month"].SortByColumn = table.Columns["Month Number"];
    table.Columns["Month Short"].SortByColumn = table.Columns["Month Number"];
    table.Columns["Day of Week"].SortByColumn = table.Columns["Day of Week Number"];
    table.Columns["Day of Week Short"].SortByColumn = table.Columns["Day of Week Number"];

    // Display folders
    table.Columns["Year"].DisplayFolder = "1. Year";

    table.Columns["Year Quarter"].DisplayFolder = "2. Quarter";
    table.Columns["Year Quarter Number"].DisplayFolder = "2. Quarter";
    table.Columns["Quarter"].DisplayFolder = "2. Quarter";

    table.Columns["Year Month"].DisplayFolder = "3. Month";
    table.Columns["Year Month Short"].DisplayFolder = "3. Month";
    table.Columns["Year Month Number"].DisplayFolder = "3. Month";
    table.Columns["Month"].DisplayFolder = "3. Month";
    table.Columns["Month Short"].DisplayFolder = "3. Month";
    table.Columns["Month Number"].DisplayFolder = "3. Month";

    table.Columns["Date"].DisplayFolder = "4. Date";

    table.Columns["Day of Week"].DisplayFolder = "5. Day";
    table.Columns["Day of Week Short"].DisplayFolder = "5. Day";
    table.Columns["Day of Week Number"].DisplayFolder = "5. Day";
    table.Columns["Working Day"].DisplayFolder = "5. Day";
    table.Columns["Working Day Number"].DisplayFolder = "5. Day";

    // Format strings
    table.Columns["Date"].FormatString = "mm/dd/yyyy";

    table.Columns["Year Quarter Number"].FormatString = "0";
    table.Columns["Year Month Number"].FormatString = "0";
    table.Columns["Month Number"].FormatString = "0";
    table.Columns["Day of Week Number"].FormatString = "0";
    table.Columns["Working Day Number"].FormatString = "0";

    // Hide columns used only for sorting
    table.Columns["Year Quarter Number"].IsHidden = true;
    table.Columns["Year Month Number"].IsHidden = true;

    // Descriptions
    table.Columns["Date"].Description = "YYYY-MM-DD";
    table.Columns["Year"].Description = "YYYY";

    table.Columns["Year Quarter"].Description = "Q1 2024";
    table.Columns["Year Quarter Number"].Description = "Used for sorting Year Quarter";
    table.Columns["Quarter"].Description = "I.e. Q1";

    table.Columns["Year Month"].Description = "I.e. January 2024";
    table.Columns["Year Month Short"].Description = "I.e. Jan 2024";
    table.Columns["Year Month Number"].Description = "Used for sorting Year Month";
    table.Columns["Month"].Description = "I.e. January";
    table.Columns["Month Short"].Description = "I.e. Jan";
    table.Columns["Month Number"].Description = "I.e. 1 (=January)";

    table.Columns["Day of Week"].Description = "I.e. Monday";
    table.Columns["Day of Week Short"].Description = "I.e. Mon";
    table.Columns["Day of Week Number"].Description = "1=Sunday to 7=Saturday";
    table.Columns["Working day"].Description = "TRUE/FALSE";
    table.Columns["Working Day Number"].Description = "Accumulated work day no from first date in calendar";

    // Table description
    table.Description = "Date-table";
}
catch
{
    Error("No selected date-table or see if column definitions are matching => No formatting");
}
