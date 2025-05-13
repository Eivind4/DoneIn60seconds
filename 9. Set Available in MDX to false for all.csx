
// Setting all columns to not available in MDX.
// Se information:  https://data-mozart.com/hidden-little-gem-that-can-save-your-power-bi-life/

// Ensure that IsAvailableInMDX is set to false if it is not to be used in Excel or as sorting for other columns etc

foreach (var table in Model.Tables)
{
    foreach (var column in table.Columns)
    {
 
        column.IsAvailableInMDX = false;
       
    }
}
