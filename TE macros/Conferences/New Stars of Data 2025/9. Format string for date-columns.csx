foreach (var table in Model.Tables)
{
    foreach (var column in table.Columns)
    {
        // Check if the column name contains "Date" (case insensitive)
        bool containsDate = column.Name.IndexOf("Date", StringComparison.OrdinalIgnoreCase) >= 0;

        // Check if the column is of DateTime data type
        bool isDateTime = column.DataType == TabularEditor.TOMWrapper.DataType.DateTime;

        // Check if the column format is NOT "mm/dd/yyyy"
        bool incorrectFormat = column.FormatString != "mm/dd/yyyy";

        // Apply the format change if all conditions are met
        if (containsDate && isDateTime && incorrectFormat)
        {
            column.FormatString = "mm/dd/yyyy";
        }
    }
}
