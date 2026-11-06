
/*
 * Title: Mark as Date-Table and create date control measures + relationship
 * 
 * Author: Eivind Haugen
 * 
 * This script, when executed, will perform necessary and generic operations for a date table. The following steps are applied:
 * 1. Select the date column and mark it as a date-table
 * 2. Create measures that later can be applied to hide future dates in time calculation, either by:
 *        2a: The minimum of the selected date in the applied filter from the calendar or today
 *        2b: Last fact date, selected date from a fact table
 * 3. Setting up relationship to a fact table, the same date used as the last fact date
 *          NB! If more fact-tables, it will have to be done manually
 * 4. Apply best practice:
 *       4a: Ensure that numeric columns are not summarized
 *       4b: Apply a format string for the date-columns MM/DD/YYYY (as this is the condition in best-practice analyzer) 
 */


using System.Drawing;
using System.Windows.Forms;
using System.Linq;

// Hide the 'Running Macro' spinbox
ScriptHelper.WaitFormVisible = false;

// Check if a table is selected in the UI
if (Selected.Table == null)
{
    Error("No table selected in the UI. Please select a table before running the script.");
    return;
}

// Use the selected table from the UI
string _CalendarTable = Selected.Table.Name;

try
{
    // Retrieve all columns in the selected date table that are explicitly formatted as dates
    var _dateTableColumns = Model.AllColumns
                                  .Where(col => col.Table.Name == _CalendarTable 
                                                && col.DataType == DataType.DateTime)
        .ToList();

    if (_dateTableColumns.Count == 0)
    {
        Error("No columns with a valid date format found in the selected table.");
        return;
    }

    try
    {
        // Select the date-key in the main date-table
        string _dateColumn = SelectColumn(_dateTableColumns, null, "Select date-key to mark as date table").DaxObjectFullName;

        // Extract the actual column name (format: 'Table'[Column])
        string _selectedColumnName = _dateColumn.Split('[', ']')[1];

        // Extract the table name from _dateColumn (format: 'Table'[Column])
        string _CalendarTableName = _dateColumn.Split('\'')[1];

        // Mark table as date table
        Model.Tables[_CalendarTable].DataCategory = "Time";
        Model.Tables[_CalendarTable].Columns[_selectedColumnName].IsKey = true;

        // Filter tables with DataCategory NOT set to "Time"
        var otherTables = Model.Tables.Where(t => t.DataCategory != "Time");

        // List of all DateTime columns in tables where DataCategory <> "Time"
        var _dateOtherTables = otherTables
            .SelectMany(t => t.Columns)
            .Where(c => c.DataType == DataType.DateTime)
            .ToList();

        if (_dateOtherTables.Count == 0)
        {
            Info("No date columns found in other tables. Proceeding without creating fact table measures.");
        }
        else
        {
            try
            {
                // Select the date for last "fact" date
                string _factDateCalendar = SelectColumn(_dateOtherTables, null, "Select date from fact-table: \n - Create measure for last date").DaxObjectFullName;
                string _FactTable = _factDateCalendar.Split('[')[0].Trim('\'');

                // Define the measure expression for max calendar date
                string measureExpression = $@"MIN(MAX({_dateColumn}), TODAY())";

                // Normalize expression (remove spaces, newlines, tabs)
                string normalizedExpression = measureExpression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

                // Check if a measure with the same DAX expression exists anywhere in the model
                bool measureExists = Model.AllMeasures.Any(m =>
                    m.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") == normalizedExpression
                );

                // Create measure only if it does not exist
                if (!measureExists)
                {
                    var measure1 = Model.Tables[_CalendarTable].AddMeasure("Max Calendar date");
                    measure1.DisplayFolder = "0. Date Control measures";
                    measure1.FormatString = "dd-MM-yyyy";
                    measure1.Description = "The last date selected from a slicer from calendar. Setting today() if last date is after todays' date.";
                    measure1.Expression = measureExpression;
                }
                else
                {
                    Info("A measure with the same expression for finding max calendar date already exists, skipping creation.");
                }

                // Define the measure expression for setting fact table date
                string transactionMeasureExpression = $@"CALCULATE (
    MAX({_factDateCalendar}),
    ALL ('{_FactTable}')
)";

                // Normalize expression (remove spaces, newlines, tabs)
                string normalizedTransactionExpression = transactionMeasureExpression
                    .Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

                // Check if a measure with the same DAX expression exists anywhere in the model
                bool transactionMeasureExists = Model.AllMeasures.Any(m =>
                    m.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") == normalizedTransactionExpression
                );

                // Create measure only if it does not exist
                if (!transactionMeasureExists)
                {
                    var measure2 = Model.Tables[_CalendarTable].AddMeasure("Last fact table date");
                    measure2.DisplayFolder = "0. Date Control measures";
                    measure2.FormatString = "dd-MM-yyyy";
                    measure2.Description = "The last fact table date.";
                    measure2.Expression = transactionMeasureExpression;
                }
                else
                {
                    Info("A measure with the same expression for finding the last fact table date already exists, skipping creation.");
                }

                // Show Yes/No dialog to confirm relationship creation
                DialogResult dialogResult = MessageBox.Show(
                    "Do you want to add a relationship to this column? (Click 'Yes' to add relationship)", 
                    "Relationship Confirmation", 
                    MessageBoxButtons.YesNo
                );

                if (dialogResult == DialogResult.Yes)
                {
                    // Creating relationship
                    // Extract column name properly
                    string _factDateColumnName = _factDateCalendar.Split('[')[1].TrimEnd(']');

                    // Check if a relationship already exists from the Date-table
                    var existingRelationship = Model.Relationships.FirstOrDefault(rel =>
                        rel.ToTable.Name == _CalendarTableName &&
                        rel.ToColumn.Name == _selectedColumnName 
                    );

                    if (existingRelationship == null)
                    {
                        var factTable = Model.Tables[_FactTable];

                        // Check if the table exists
                        var calendarTable = Model.Tables[_CalendarTableName];
                        if (calendarTable == null)
                        {
                            Error($"Table {_CalendarTableName} does not exist.");
                            return;
                        }

                        // Get the to-column (calendar table column)
                        var toColumn = calendarTable.Columns[_selectedColumnName];

                        if (factTable.Columns.Contains(_factDateColumnName))
                        {
                            var fromColumn = factTable.Columns[_factDateColumnName];

                            if (toColumn != null && fromColumn != null)
                            {
                                try
                                {
                                    var singleColumnRelationship = Model.AddRelationship();
                                    singleColumnRelationship.FromColumn = fromColumn;
                                    singleColumnRelationship.ToColumn = toColumn;
                                    singleColumnRelationship.FromCardinality = RelationshipEndCardinality.Many;
                                    singleColumnRelationship.ToCardinality = RelationshipEndCardinality.One;
                                    singleColumnRelationship.IsActive = true;

                                    Info("Relationship successfully created!");
                                }
                                catch (Exception ex)
                                {
                                    Error($"Error creating relationship: {ex.Message}");
                                }
                            }
                            else
                            {
                                Error("Could not find valid columns to create the relationship.");
                            }
                        }
                        else
                        {
                            Error($"Column {_factDateColumnName} does not exist in table {_FactTable}.");
                        }
                    }
                    else
                    {
                        Info("Relationship already exists, skipping creation.");
                    }
                }
                else
                {
                    Info("Skipping the step for creating a relationship.");
                }
            }
            catch (Exception ex)
            {
                Info($"No date selected for last fact date or error occurred: {ex.Message}");
            }
        }

        // Set Table Group
        Model.Tables[_CalendarTable].TableGroup = "Dim Calendar";

        // Apply Best Practice Rules
        // Hide foreign key columns & prevent numeric columns from summarizing
        foreach (var table in Model.Tables)
        {
            if (table.Name == _CalendarTable)
            {
                foreach (var column in table.Columns)
                {
                    if (column.DataType == DataType.Int64)
                    {
                        column.SummarizeBy = AggregateFunction.None;
                    }
                }
            }
        }

        // Format all date columns
        foreach (var table in Model.Tables)
        {
            foreach (var column in table.Columns)
            {
                if (column.DataType == DataType.DateTime)
                {
                    column.FormatString = "MM/dd/yyyy";
                }
            }
        }

        Info("Date table setup completed successfully!");
    }
    catch (Exception ex)
    {
        Error($"Error selecting date-key column: {ex.Message}");
    }
}
catch (Exception ex)
{
    Error($"Error processing the selected table: {ex.Message}");
}