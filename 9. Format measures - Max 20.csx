if (Selected.Measures.Count == 0)
{
    Error("No measures selected. Please select one or more measures to format.");
    return;
}

if (Selected.Measures.Count > 20)
{
    Error("Too many measures selected. Reduce usage of DAXformatter");
    return;
}

foreach (var m in Selected.Measures)
{
    m.Expression = FormatDax(m.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
    /* Format only the selected measures using DAX Formatter */
}
