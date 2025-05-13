
 // Title: Initial formatting of Date-Table
 // 
 // Author: Eivind Haugen
 // 
 // This script, when executed, will perform necessary and generic operations for a date table. The following steps are applied:
 // 1. Select the date column and mark it as a date-table
 // 2. Create measures that later can be applied to hide future dates in time calculation, either by:
 //        2a: The minimum of the selected date in the applied filter from the calendar or today
 //        2b: Last fact date, selected date from a fact table
 // 3. Setting up relationship to a fact table, the same date used as the last fact date
 //          NB! If more fact-tables, it will have to be done manually
 // 4. Apply best practice:
 //       4a: Ensure that numeric columns are not summarized
 //       4b: Apply a format string for the date-columns MM/DD/YYYY (as this is the condition in best-practice analyzer) 
 



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

        // Apply best practices
        foreach (var column in Model.Tables[_CalendarTable].Columns)
        {
            if (column.DataType == DataType.Int64)
                column.SummarizeBy = AggregateFunction.None;
        }

        // Create Relationship (only if not exists)
        string _factDateColumnName = _factDateCalendar.Split('[')[1].TrimEnd(']');
        var existingRelationship = Model.Relationships.FirstOrDefault(rel =>
            rel.ToTable.Name == _CalendarTableName &&
            rel.ToColumn.Name == _selectedColumnName
        );

if (existingRelationship == null)
{
    string message = string.Format(
        "No existing relationship found between '{0}'[{1}] and '{2}'[{3}].\n\nDo you want to create it?",
        _FactTable, _factDateColumnName, _CalendarTableName, _selectedColumnName
    );

    var result = MessageBox.Show(
        message,
        "Create Relationship?",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    );

    if (result == DialogResult.Yes)
    {
        var calendarTable = Model.Tables[_CalendarTableName];
        var factTable = Model.Tables[_FactTable];
        var fromColumn = factTable.Columns[_factDateColumnName];
        var toColumn = calendarTable.Columns[_selectedColumnName];

        if (fromColumn != null && toColumn != null)
        {
            var rel = Model.AddRelationship();
            rel.FromColumn = fromColumn;
            rel.ToColumn = toColumn;
            rel.FromCardinality = RelationshipEndCardinality.Many;
            rel.ToCardinality = RelationshipEndCardinality.One;
            rel.IsActive = true;

            Output(string.Format(
                "Relationship created between '{0}'[{1}] and '{2}'[{3}].",
                _FactTable, _factDateColumnName, _CalendarTableName, _selectedColumnName
            ));
        }
        else
        {
            Error("Could not find valid columns to create the relationship.");
        }
    }
    else
    {
        Output("User chose not to create the relationship.");
    }
}
else
{
    Output("Relationship already exists between calendar and fact table.");
}

    }
    catch
    {
        Error("No date-key selected to mark as date table.");
    }
