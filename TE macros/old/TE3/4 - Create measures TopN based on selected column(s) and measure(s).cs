
// Author: Eivind Haugen

// 1. If Top N does not exist it will be created
// 2. Need to select both column and measure. All combinations will be created if multiple selections
// 3. It will create rankings based on the measure selected for the given column


// Check if the "Top N" table already exists
if (Model.Tables.Contains("Top N") == false)
{
    // Create the calculated table since it doesn't already exist
    var calculatedTable1 = Model.AddCalculatedTable();
    calculatedTable1.TableGroup = "Parameters";
    calculatedTable1.Name = "Top N";
    Model.Tables["Top N"].Partitions["New Calculated Table"].Expression = "GENERATESERIES(1, 20, 1)";
var measure1 = Model.Tables["Top N"].AddMeasure("Top N Value");
measure1.Expression = "SELECTEDVALUE('Top N'[Value])";
measure1.Description = "Selected parameter value";
measure1.FormatString = "#,0";

calculatedTable1.Columns["Value"].SummarizeBy = AggregateFunction.None;
calculatedTable1.Columns["Value"].Description = "Selected parameter for Top N";
calculatedTable1.Columns["Value"].FormatString = "#,0";

}
else
{

Info("Table 'Top N' already exists. Skipping creation.");
}



foreach (var c in Selected.Columns)
{
    // Define the selected table and column names
    string selectedTable = c.Table.Name;    // Access the Name property of the Table object
    string selectedColumn = c.Name;         // Get the column name

    foreach (var m in Selected.Measures)
    {
        // Define the selected measure name
        string selectedMeasure = m.Name;   // Use the Name property of the measure
        string selectedMeasureTable = m.Table.Name; // Use the Table property of the measure

        // Build the measure expression dynamically, replacing placeholders with selected table, column, and measure names
        string measureExpression = $@"

        // Reference: https://www.sqlbi.com/articles/introducing-rankx-in-dax/

        VAR _ValuesToRank = [Top N Value]

        VAR _Result =
        IF (
            ISINSCOPE ( '{selectedTable}'[{selectedColumn}] ),
            
                    VAR _RankingValue =
                        RANKX (
                            ALLSELECTED ( '{selectedTable}'[{selectedColumn}] ),
                            [{selectedMeasure}],
                            ,
                            ,
                            SKIP
                        )
                    RETURN
                        IF (
                            _RankingValue > 0 && _RankingValue <= _ValuesToRank,
                            _RankingValue
                        )
                )
        
        RETURN
            _Result
        ";

        // Add the measure to the model with the dynamically created expression
        var measureName = $"Ranking Top for {selectedColumn} by {selectedMeasure}";
        var measure1 = Model.Tables[selectedMeasureTable].AddMeasure(measureName);
        measure1.Expression = measureExpression;
        // Add format and display folder
        measure1.FormatString = "#,0";
        measure1.DisplayFolder  = "Ranking";
        measure1.Description = "Top N ranking for {selectedColumn} by {selectedMeasure}";
        
        
         // Apply short format to DAX expression
            measure1.Expression = FormatDax(measure1.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
    }
}
