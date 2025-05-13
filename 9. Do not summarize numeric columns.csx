foreach (var table in Model.Tables)
{
    foreach (var column in table.Columns)
    {
        // Set SummarizeBy to None for columns ending in "Key" or "ID"
        if (column.Name.EndsWith("Key") || column.Name.EndsWith("ID"))
        {
            column.SummarizeBy = AggregateFunction.None;
        }

        // Set SummarizeBy to None for numeric columns not already set to None, not hidden, and not from hidden tables
        if ((column.DataType == DataType.Int64 ||
             column.DataType == DataType.Decimal ||
             column.DataType == DataType.Double) &&
            column.SummarizeBy != AggregateFunction.None &&
            !column.IsHidden &&
            !column.Table.IsHidden)
        {
            column.SummarizeBy = AggregateFunction.None;
        }
    }
}
