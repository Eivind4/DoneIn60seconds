using System.Linq;
using System.Windows.Forms;

ScriptHelper.WaitFormVisible = false;

Func<string, string> Normalize = s => s
    .Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

// ── Helper: simple text input dialog ──────────────────────────────
Func<string, string, string> InputBox = (prompt, defaultValue) =>
{
    var inputForm = new Form();
    inputForm.Text = prompt;
    inputForm.Width = 380;
    inputForm.Height = 150;
    inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
    inputForm.StartPosition = FormStartPosition.CenterScreen;
    inputForm.MaximizeBox = false;

    var label     = new Label()   { Left = 20, Top = 15, Width = 320, Text = prompt };
    var textBox   = new TextBox() { Left = 20, Top = 40, Width = 320, Text = defaultValue };
    var btnOk     = new Button()  { Text = "OK",     Left = 180, Top = 75, Width = 80, DialogResult = DialogResult.OK };
    var btnCancel = new Button()  { Text = "Cancel", Left = 270, Top = 75, Width = 80, DialogResult = DialogResult.Cancel };

    inputForm.Controls.AddRange(new Control[] { label, textBox, btnOk, btnCancel });
    inputForm.AcceptButton = btnOk;
    inputForm.CancelButton = btnCancel;

    return inputForm.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : null;
};

// ── Step 1: Checkbox dialog to select which UDFs to create ────────
var form = new Form();
form.Text = "Select UDFs to create";
form.Width = 380;
form.Height = 230;
form.FormBorderStyle = FormBorderStyle.FixedDialog;
form.StartPosition = FormStartPosition.CenterScreen;
form.MaximizeBox = false;

var chk5a = new CheckBox() { Text = "5a: Filter by last fact table date",    Left = 20, Top = 20, Width = 320, Checked = true };
var chk5b = new CheckBox() { Text = "5b: Filter by last selected date",      Left = 20, Top = 55, Width = 320, Checked = true };
var chk5c = new CheckBox() { Text = "5c: Filter by history flag column",     Left = 20, Top = 90, Width = 320, Checked = true };

var btnOk2     = new Button() { Text = "OK",     Left = 180, Top = 140, Width = 80, DialogResult = DialogResult.OK };
var btnCancel2 = new Button() { Text = "Cancel", Left = 270, Top = 140, Width = 80, DialogResult = DialogResult.Cancel };

form.Controls.AddRange(new Control[] { chk5a, chk5b, chk5c, btnOk2, btnCancel2 });
form.AcceptButton = btnOk2;
form.CancelButton = btnCancel2;

if (form.ShowDialog() != DialogResult.OK)
{
    Info("Cancelled.");
    return;
}

bool create5a = chk5a.Checked;
bool create5b = chk5b.Checked;
bool create5c = chk5c.Checked;

if (!create5a && !create5b && !create5c)
{
    Info("Nothing selected, exiting.");
    return;
}

// ── Step 2: Pick the date column from the calendar table ──────────
string _dateColumn    = null;
string _CalendarTable = null;

if (create5b || create5c)
{
    var dateColumns = Model.AllColumns
        .Where(col => col.DataType == DataType.DateTime
                   && col.Table.DataCategory == "Time")
        .ToList();

    if (dateColumns.Count == 0)
    {
        Error("No DateTime columns found in a table marked as 'Time'.");
        return;
    }

    var selectedDateCol = SelectColumn(dateColumns, null, "Select the date column from the calendar table");
    _dateColumn    = selectedDateCol.DaxObjectFullName;
    _CalendarTable = selectedDateCol.Table.Name;
}

// ── Helper to create a UDF if name and expression are both new ────
Action<string, string, string> CreateUDF = (fnName, fnExpression, fnDescription) =>
{
    bool nameExists       = Model.Functions.Any(f => f.Name == fnName);
    bool expressionExists = Model.Functions.Any(f => Normalize(f.Expression) == Normalize(fnExpression));

    if (nameExists || expressionExists)
    {
        Info($"Skipping '{fnName}': a function with the same name or expression already exists.");
        return;
    }

    var fn = Model.AddFunction(fnName);
    fn.Expression  = fnExpression;
    fn.Description = fnDescription;
    fn.FormatDax();
};

// ── 5a: Filter by last fact table date ───────────────────────────
if (create5a)
{
    var nonCalendarDateColumns = Model.AllColumns
        .Where(col => col.DataType == DataType.DateTime 
                   && col.Table.DataCategory != "Time")
        .ToList();

    if (nonCalendarDateColumns.Count == 0)
    {
        Info("No DateTime columns found outside the calendar table, skipping 5a.");
    }
    else
    {
        var factDateCol   = SelectColumn(nonCalendarDateColumns, null, "5a: Select the date column from the fact table");
        string _factDate  = factDateCol.DaxObjectFullName;
        string _factTable = factDateCol.Table.Name;

        CreateUDF(
            "Local.TimeIntelligence.Filter.HistoricalByLastFactDate",
            $@"( 
//https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    valueExpr : ANYREF EXPR
) =>
VAR LastDayAvailable =
    CALCULATE (
        MAX ( {_factDate} ),
        ALL ( '{_factTable}' )
    )
RETURN
CALCULATE (
    valueExpr,
    KEEPFILTERS (
        {_dateColumn} <= LastDayAvailable
    )
)",
            $"Filters historical dates up to the last available date in '{_factTable}' ({_factDate}), ignoring all filters on that table."
        );
    }
}

// ── 5b: Filter by last selected calendar date ─────────────────────
if (create5b)
{
    CreateUDF(
        "Local.TimeIntelligence.Filter.HistoricalByMaxCalendarDate",
        $@"( 
    // https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    // Expression to evaluate historical dates
    valueExpr : ANYREF EXPR
) =>
CALCULATE (
    valueExpr,
    KEEPFILTERS (
        {_dateColumn} <= MAX ( {_dateColumn} )
    )
)",
        "Filters historical dates up to the last date selected in the current filter context (e.g. from a slicer)."
    );
}

// ── 5c: Filter by history flag column ────────────────────────────
if (create5c)
{
    var historyFlagCol = SelectColumn(Model.AllColumns.ToList(), null, "5c: Select the history flag column (e.g. is_History)");
    string _historyCol = historyFlagCol.DaxObjectFullName;

    string _filterValue = InputBox("Enter the filter value for the history flag (e.g. 1 or Yes)", "1");

    if (string.IsNullOrWhiteSpace(_filterValue))
    {
        Info("No filter value entered, skipping 5c.");
    }
    else
    {
        string _filterValueFormatted = double.TryParse(_filterValue, out _)
            ? _filterValue
            : $"\"{_filterValue}\"";

        CreateUDF(
            "Local.TimeIntelligence.Filter.HistoricalDatesColumn",
            $@"( 
    // https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
    // Expression to evaluate historical dates
    valueExpr : ANYREF EXPR
) =>
CALCULATE (
    valueExpr,
    CALCULATETABLE (
        VALUES ( {_dateColumn} ),
        {_historyCol} = {_filterValueFormatted}
    )
)",
            $"Filters historical dates where {_historyCol} = {_filterValueFormatted}."
        );
    }
}

//Info("Done — DAX UDFs created successfully!");