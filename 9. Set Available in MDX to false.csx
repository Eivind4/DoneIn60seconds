
// Ensure that IsAvailableInMDX is set to false if it is not to be used in Excel or as sorting for other columns etc

foreach (var table in Model.Tables)
{
    foreach (var column in table.Columns)
    {
        // Set IsAvailableInMDX to false for columns ending in "Key" or "ID"
        if (column.Name.EndsWith("Key") || column.Name.EndsWith("ID"))
        {
            column.IsAvailableInMDX = false;
        }

        // Set IsAvailableInMDX to false for hidden/unnecessary columns
        if (column.IsAvailableInMDX &&
            (column.IsHidden || column.Table.IsHidden) &&
            !column.UsedInSortBy.Any() &&
            !column.UsedInHierarchies.Any() &&
            !column.UsedInVariations.Any() &&
            column.SortByColumn == null)
        {
            column.IsAvailableInMDX = false;
        }
    }
}
