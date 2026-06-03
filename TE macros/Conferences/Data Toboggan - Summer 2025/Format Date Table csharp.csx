// Title: Organize Contoso Date-Table (with two columns added)
 // 
 // Author: Eivind Haugen
 //
 // Context: Table
 // 
 // This script, when executed, will perform necessary and generic operations for a date table. The following steps are applied:
 // 1. Select the date column and mark it as a date-table
 // 2. Create measures that later can be applied to hide future dates in time calculation, either by:
 //        2a: The minimum of the selected date in the applied filter from the calendar or today
 //        2b: Last fact date, selected date from a fact table
 // 3. Apply best practice:
 //       Apply a format string for the date-columns MM/DD/YYYY (as this is the condition in best-practice analyzer) 
 
 


using System.Windows.Forms;
ScriptHelper.WaitFormVisible = false;

// Use the selected table from the UI
if (Selected.Table == null)
{
    Error("No table selected in the UI. Please select a table before running the script.");
    return;
}

string _CalendarTable = Selected.Table.Name;
    try
    {
        // Select a column to mark as date key
        var _dateTableColumns = Model.AllColumns
            .Where(col => col.Table.Name == _CalendarTable && col.DataType == DataType.DateTime)
            .ToList();

        if (_dateTableColumns.Count == 0)
        {
            Error("No columns with a valid date format found in the selected table.");
            return;
        }

        string _dateColumn = SelectColumn(_dateTableColumns, null, "Select date-key to mark as date table").DaxObjectFullName;

        string _selectedColumnName = _dateColumn.Split('[', ']')[1];
        string _CalendarTableName = _dateColumn.Split('\'')[1];

        Model.Tables[_CalendarTable].DataCategory = "Time";
        Model.Tables[_CalendarTable].Columns[_selectedColumnName].IsKey = true;

        // Select date column from a fact table
        var otherTables = Model.Tables.Where(t => t.DataCategory != "Time");
        var _dateOtherTables = otherTables
            .SelectMany(t => t.Columns)
            .Where(c => c.DataType == DataType.DateTime)
            .ToList();

            string _factDateCalendar = SelectColumn(_dateOtherTables, null, "Select date from fact-table to Create Measure for the Last date").DaxObjectFullName;
        string _FactTable = _factDateCalendar.Split('[')[0].Trim('\'');

        // Measure 1: Max Calendar Date
        string measureExpression = string.Format("MIN(MAX({0}), TODAY())", _dateColumn);
        string normalizedExpression = measureExpression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

        bool measureExists = Model.AllMeasures.Any(m =>
            m.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") == normalizedExpression
        );

        if (!measureExists)
        {
            var measure1 = Model.Tables[_CalendarTable].AddMeasure("Max Calendar date");
            measure1.DisplayFolder = "0. Date Control measures";
            measure1.FormatString = "dd-MM-yyyy";
            measure1.Description = "The last date selected from a slicer from calendar. Setting today() if last date is after today's date.";
            measure1.Expression = measureExpression;
        }

        // Measure 2: Last Fact Table Date
        string transactionMeasureExpression = string.Format("CALCULATE(MAX({0}), ALL ('{1}'))", _factDateCalendar, _FactTable);
        string normalizedTransactionExpression = transactionMeasureExpression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

        bool transactionMeasureExists = Model.AllMeasures.Any(m =>
            m.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") == normalizedTransactionExpression
        );

        if (!transactionMeasureExists)
        {
            var measure2 = Model.Tables[_CalendarTable].AddMeasure("Last fact table date");
            measure2.DisplayFolder = "0. Date Control measures";
            measure2.FormatString = "dd-MM-yyyy";
            measure2.Description = "The last fact table date.";
            measure2.Expression = transactionMeasureExpression;
        }



        // Apply column formatting, sorting, and display folders
        try
        {
            var table = Model.Tables[_CalendarTable];

            // Sorting (only set if columns exist)
            if (table.Columns.Any(c => c.Name == "Year Quarter") && table.Columns.Any(c => c.Name == "Year Quarter Number"))
                table.Columns["Year Quarter"].SortByColumn = table.Columns["Year Quarter Number"];
            
            if (table.Columns.Any(c => c.Name == "Year Month Short") && table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Short"].SortByColumn = table.Columns["Year Month Number"];
            
            if (table.Columns.Any(c => c.Name == "Year Month") && table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month"].SortByColumn = table.Columns["Year Month Number"];
            
            if (table.Columns.Any(c => c.Name == "Month") && table.Columns.Any(c => c.Name == "Month Number"))
                table.Columns["Month"].SortByColumn = table.Columns["Month Number"];
            
            if (table.Columns.Any(c => c.Name == "Month Short") && table.Columns.Any(c => c.Name == "Month Number"))
                table.Columns["Month Short"].SortByColumn = table.Columns["Month Number"];
            
            if (table.Columns.Any(c => c.Name == "Day of Week") && table.Columns.Any(c => c.Name == "Day of Week Number"))
                table.Columns["Day of Week"].SortByColumn = table.Columns["Day of Week Number"];
            
            if (table.Columns.Any(c => c.Name == "Day of Week Short") && table.Columns.Any(c => c.Name == "Day of Week Number"))
                table.Columns["Day of Week Short"].SortByColumn = table.Columns["Day of Week Number"];
            
              if (table.Columns.Any(c => c.Name == "Year Month Slicer") && table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Slicer"].SortByColumn = table.Columns["Year Month Number"];

            // Display folders
            if (table.Columns.Any(c => c.Name == "Year"))
                table.Columns["Year"].DisplayFolder = "1. Year";

            if (table.Columns.Any(c => c.Name == "Year Quarter"))
                table.Columns["Year Quarter"].DisplayFolder = "2. Quarter";
            if (table.Columns.Any(c => c.Name == "Year Quarter Number"))
                table.Columns["Year Quarter Number"].DisplayFolder = "2. Quarter";
            if (table.Columns.Any(c => c.Name == "Quarter"))
                table.Columns["Quarter"].DisplayFolder = "2. Quarter";

            if (table.Columns.Any(c => c.Name == "Year Month"))
                table.Columns["Year Month"].DisplayFolder = "3. Month";
            if (table.Columns.Any(c => c.Name == "Year Month Short"))
                table.Columns["Year Month Short"].DisplayFolder = "3. Month";
            if (table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Number"].DisplayFolder = "3. Month";
            if (table.Columns.Any(c => c.Name == "Month"))
                table.Columns["Month"].DisplayFolder = "3. Month";
            if (table.Columns.Any(c => c.Name == "Month Short"))
                table.Columns["Month Short"].DisplayFolder = "3. Month";
            if (table.Columns.Any(c => c.Name == "Month Number"))
                table.Columns["Month Number"].DisplayFolder = "3. Month";

            if (table.Columns.Any(c => c.Name == "Date"))
                table.Columns["Date"].DisplayFolder = "4. Date";

            if (table.Columns.Any(c => c.Name == "Day of Week"))
                table.Columns["Day of Week"].DisplayFolder = "5. Day";
            if (table.Columns.Any(c => c.Name == "Day of Week Short"))
                table.Columns["Day of Week Short"].DisplayFolder = "5. Day";
            if (table.Columns.Any(c => c.Name == "Day of Week Number"))
                table.Columns["Day of Week Number"].DisplayFolder = "5. Day";
            if (table.Columns.Any(c => c.Name == "Working Day"))
                table.Columns["Working Day"].DisplayFolder = "5. Day";
            if (table.Columns.Any(c => c.Name == "Working Day Number"))
                table.Columns["Working Day Number"].DisplayFolder = "5. Day";
              if (table.Columns.Any(c => c.Name == "Year Month Slicer"))
                  table.Columns["Year Month Slicer"].DisplayFolder = "8. Slicer";
              if (table.Columns.Any(c => c.Name == "is_History"))
                  table.Columns["is_History"].DisplayFolder = "6. Boolean";

              // Format strings for dates in date-table
            if (table.Columns.Any(c => c.Name == "Date"))
                table.Columns["Date"].FormatString = "mm/dd/yyyy";

            if (table.Columns.Any(c => c.Name == "Year Quarter Number"))
                table.Columns["Year Quarter Number"].FormatString = "0";
            if (table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Number"].FormatString = "0";
            if (table.Columns.Any(c => c.Name == "Month Number"))
                table.Columns["Month Number"].FormatString = "0";
            if (table.Columns.Any(c => c.Name == "Day of Week Number"))
                table.Columns["Day of Week Number"].FormatString = "0";
            if (table.Columns.Any(c => c.Name == "Working Day Number"))
                table.Columns["Working Day Number"].FormatString = "0";

            // Hide columns used only for sorting
            if (table.Columns.Any(c => c.Name == "Year Quarter Number"))
                table.Columns["Year Quarter Number"].IsHidden = true;
            if (table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Number"].IsHidden = true;

            // Descriptions
            if (table.Columns.Any(c => c.Name == "Date"))
                table.Columns["Date"].Description = "YYYY-MM-DD";
            if (table.Columns.Any(c => c.Name == "Year"))
                table.Columns["Year"].Description = "YYYY";

            if (table.Columns.Any(c => c.Name == "Year Quarter"))
                table.Columns["Year Quarter"].Description = "Q1 2024";
            if (table.Columns.Any(c => c.Name == "Year Quarter Number"))
                table.Columns["Year Quarter Number"].Description = "Used for sorting Year Quarter";
            if (table.Columns.Any(c => c.Name == "Quarter"))
                table.Columns["Quarter"].Description = "I.e. Q1";

            if (table.Columns.Any(c => c.Name == "Year Month"))
                table.Columns["Year Month"].Description = "I.e. January 2024";
            if (table.Columns.Any(c => c.Name == "Year Month Short"))
                table.Columns["Year Month Short"].Description = "I.e. Jan 2024";
            if (table.Columns.Any(c => c.Name == "Year Month Number"))
                table.Columns["Year Month Number"].Description = "Used for sorting Year Month";
            if (table.Columns.Any(c => c.Name == "Month"))
                table.Columns["Month"].Description = "I.e. January";
            if (table.Columns.Any(c => c.Name == "Month Short"))
                table.Columns["Month Short"].Description = "I.e. Jan";
            if (table.Columns.Any(c => c.Name == "Month Number"))
                table.Columns["Month Number"].Description = "I.e. 1 (=January)";

            if (table.Columns.Any(c => c.Name == "Day of Week"))
                table.Columns["Day of Week"].Description = "I.e. Monday";
            if (table.Columns.Any(c => c.Name == "Day of Week Short"))
                table.Columns["Day of Week Short"].Description = "I.e. Mon";
            if (table.Columns.Any(c => c.Name == "Day of Week Number"))
                table.Columns["Day of Week Number"].Description = "0=Monday to 6=Sunday";
            if (table.Columns.Any(c => c.Name == "Working Day"))
                table.Columns["Working Day"].Description = "TRUE/FALSE";
                if (table.Columns.Any(c => c.Name == "is_History"))
                    table.Columns["is_History"].Description = "Before today's date";
                  if (table.Columns.Any(c => c.Name == "Year Month Slicer"))
                      table.Columns["Year Month Slicer"].Description = "Current month if month and year is the same as today";
            if (table.Columns.Any(c => c.Name == "Working Day Number"))
                table.Columns["Working Day Number"].Description = "Accumulated work day no from first date in calendar";

            // Table description
            table.Description = "Date-table";
        }
        catch
        {
            Output("Some column formatting could not be applied - check if column names match expected format.");
        }




    }
    catch
    {
        Error("No date-key selected to mark as date table.");
    }
