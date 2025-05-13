
// This script includes the select time calculations that are to be included. The idea is from the following script, but here it is included as a Calculation Group instead and with some modifications
//        Data-Goblin inspiration: https://github.com/data-goblin/powerbi-macguyver-toolbox/blob/main/tabular-editor-scripts/csharp-scripts/add-dax-templates/add-time-intelligence.csx


// There are some input that are dynamically included in the generation of time intelligence calculations:
//     1. Date column to be used from the date table (needs to be created first)
//     2. Measure that shows the last data from calendar or transaction (needs to be calculated first and have the correct formatString containing "yy"). The purpose is to hide future dates in calculations, where applicable
//     3. Based on the selected Time intelligence calculations, it is logic on which input is required. 

// Naming is aiming at following this pattern: https://www.daxpatterns.com/standard-time-related-calculations/
// Referance to source for the expressions are included in the measure expressions

// If you know of some measures with % format, this can be updated in the logic for the calculation item before running the script



#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;


// Don't show the script execution dialog or cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;


// ─────────────────────────────────────────────────────────────
// Rename Calculation Group Table and Column if it's a new one

// Define inline input dialog function 
Func<string, string, string, string> ShowInputDialog = (text, caption, defaultValue) =>
{
    Form prompt = new Form()
    {
        Width = 600,
        Height = 150,
        FormBorderStyle = FormBorderStyle.FixedDialog,
        Text = caption,
        StartPosition = FormStartPosition.CenterScreen
    };
    Label textLabel = new Label() { Left = 50, Top = 20, Text = text, AutoSize = true };
    TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 500, Text = defaultValue };
    Button confirmation = new Button() { Text = "OK", Left = 450, Width = 100, Top = 70, DialogResult = DialogResult.OK };
    confirmation.Click += (sender, e) => { prompt.Close(); };
    prompt.Controls.Add(textLabel);
    prompt.Controls.Add(textBox);
    prompt.Controls.Add(confirmation);
    prompt.AcceptButton = confirmation;

    return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
};

// Get the selected calculation group table. Checks the name if it is a new or existing calculation group
var cgTable = Selected.Tables.First() as CalculationGroupTable;

if (cgTable != null && cgTable.Name.StartsWith("New Calculation Group"))
{
    string defaultTableName = "Time Intelligence";
    string defaultColumnName = "Time Calculation";

    string newTableName = ShowInputDialog(
        "Enter new name for the calculation group table:",
        "Rename Calculation Group Table",
        defaultTableName
    );

    string newColumnName = ShowInputDialog(
        "Enter name for the calculation group column:",
        "Set Calculation Group Column Name",
        defaultColumnName
    );

    if (!string.IsNullOrWhiteSpace(newTableName))
    {
        cgTable.Name = newTableName;
        cgTable.Description = "Calculation group for dynamic time intelligence.";
        cgTable.CalculationGroup.Precedence = 0;

if (!string.IsNullOrWhiteSpace(newColumnName))
{
    cgTable.Columns["Name"].Name = newColumnName;
}

        // Info(string.Format("Calculation Group renamed to '{0}' with column '{1}'", newTableName, newColumnName));
    }
    else
    {
        Warning("No table name entered. Calculation group name remains unchanged.");
    }
}




// ─────────────────────────────────────────────────────────────
// Helper: Column selector
Func<IEnumerable<Column>, string, Column> SelectColumnFromList = (columns, title) =>
{
    var form = new Form();
    var listBox = new ListBox();
    var okButton = new Button();
    var cancelButton = new Button();
    var label = new Label();

    form.Text = "Select Column"; // Keep title short (multiline not supported)
    form.Width = 400;
    form.Height = 470;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.StartPosition = FormStartPosition.CenterScreen;

    label.Text = title;  // Use the existing title input as label text
    label.Left = 10;
    label.Top = 10;
    label.Width = 360;
    label.Height = 40;
    label.AutoSize = false;

    listBox.Width = 360;
    listBox.Height = 280;
    listBox.Left = 10;
    listBox.Top = 60;
    listBox.SelectionMode = SelectionMode.One;
    listBox.DataSource = columns.Select(c => c.DaxObjectFullName).ToList();

    okButton.Text = "OK";
    okButton.Left = 220;
    okButton.Top = 360;
    okButton.DialogResult = DialogResult.OK;

    cancelButton.Text = "Cancel";
    cancelButton.Left = 300;
    cancelButton.Top = 360;
    cancelButton.DialogResult = DialogResult.Cancel;

    form.Controls.Add(label);
    form.Controls.Add(listBox);
    form.Controls.Add(okButton);
    form.Controls.Add(cancelButton);
    form.AcceptButton = okButton;
    form.CancelButton = cancelButton;

    var result = form.ShowDialog();
    if (result == DialogResult.OK && listBox.SelectedIndex >= 0)
    {
        string selectedName = listBox.SelectedItem.ToString();
        return columns.FirstOrDefault(c => c.DaxObjectFullName == selectedName);
    }
    return null;
};


// ─────────────────────────────────────────────────────────────
// Helper: Measure selector
Func<IEnumerable<Measure>, string, Measure> SelectMeasureFromList = (measures, title) =>
{
    var form = new Form();
    var listBox = new ListBox();
    var okButton = new Button();
    var cancelButton = new Button();
    var label = new Label();

    form.Text = "Select Measure"; // Keep title short
    form.Width = 400;
    form.Height = 470;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.StartPosition = FormStartPosition.CenterScreen;

    label.Text = title;  // Use the existing title input as label text
    label.Left = 10;
    label.Top = 10;
    label.Width = 360;
    label.Height = 40;
    label.AutoSize = false;

    listBox.Width = 360;
    listBox.Height = 280;
    listBox.Left = 10;
    listBox.Top = 60;
    listBox.SelectionMode = SelectionMode.One;
    listBox.DataSource = measures.Select(m => m.DaxObjectFullName).ToList();

    okButton.Text = "OK";
    okButton.Left = 220;
    okButton.Top = 360;
    okButton.DialogResult = DialogResult.OK;

    cancelButton.Text = "Cancel";
    cancelButton.Left = 300;
    cancelButton.Top = 360;
    cancelButton.DialogResult = DialogResult.Cancel;

    form.Controls.Add(label);
    form.Controls.Add(listBox);
    form.Controls.Add(okButton);
    form.Controls.Add(cancelButton);
    form.AcceptButton = okButton;
    form.CancelButton = cancelButton;

    var result = form.ShowDialog();
    if (result == DialogResult.OK && listBox.SelectedIndex >= 0)
    {
        string selectedName = listBox.SelectedItem.ToString();
        return measures.FirstOrDefault(m => m.DaxObjectFullName == selectedName);
    }
    return null;
};

// ─────────────────────────────────────────────────────────────
// Time Intelligence Selection UI

string[] calcTypes = new string[] {
    "Actual", "MTD", "QTD", "YTD", "R12", "Rolling average", "Rolling total"
};

Form tiForm = new Form();
ListBox tiListBox = new ListBox();
Button tiButton = new Button();
Label tiLabel = new Label();

tiForm.Text = "Select Time Intelligence Types";
tiForm.Width = 350;
tiForm.Height = 400;

tiLabel.Text = "Select one or more calculation types:";
tiLabel.Location = new Point(20, 10);
tiLabel.Width = 300;

tiListBox.Items.AddRange(calcTypes);
tiListBox.SelectionMode = SelectionMode.MultiExtended;
tiListBox.Location = new Point(20, 40);
tiListBox.Width = 290;
tiListBox.Height = 250;

tiButton.Text = "OK";
tiButton.Location = new Point(120, 310);
tiButton.Width = 100;
tiButton.Click += (sender, e) => { tiForm.Close(); };

tiForm.Controls.Add(tiListBox);
tiForm.Controls.Add(tiButton);
tiForm.Controls.Add(tiLabel);
tiForm.ShowDialog();

List<string> selectedValues = tiListBox.SelectedItems.Cast<string>().ToList();




var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime &&
    c.IsKey == true).ToList();
    var selectedDateCol = SelectColumnFromList(
    _dateColumns,
    "Select the DATE column for time calculations.\n\nA Date table marked as a date table is required."
);

if (selectedDateCol == null) return;
string _CalendarDate = selectedDateCol.DaxObjectFullName;
var _calendarTableName = _CalendarDate.Split('[')[0].Trim('\'');

string _LastDateAvailable = null;

bool requiresFutureDateLogic = selectedValues.Any(x => new[] { "Actual", "MTD", "QTD", "YTD", "R12", "Rolling average", "Rolling total" }.Contains(x));
if (requiresFutureDateLogic)
{
    var _dateFormattedMeasures = Model.AllMeasures
        .Where(m => m.FormatString != null && (m.FormatString.Contains("yy") || m.FormatString.Contains("yyyy")))
        .ToList();

        var selectedMeasure = SelectMeasureFromList(_dateFormattedMeasures, "Select MEASURE for hiding future dates \n\n (Cancel will create logic for last date)");
    _LastDateAvailable = selectedMeasure != null
        ? selectedMeasure.DaxObjectName
        : string.Format("MIN(MAX({0}), TODAY())", _CalendarDate);
}

string _selectedAverage = null;
if (selectedValues.Contains("Rolling average"))
{
    var _selectedTableName = _CalendarDate.Split('[')[0].Trim('\'');
    var _selectedDateTable = Model.Tables[_selectedTableName];
    var _columnsInDateTable = _selectedDateTable.Columns.ToList();

    var avgCol = SelectColumnFromList(_columnsInDateTable, "Select column for Rolling 12M Average \n\n(e.g., daily/monthly)");
    if (avgCol != null)
    {
        _selectedAverage = avgCol.DaxObjectFullName;
    }
}

// Optional debug
//string debugMsg = "Selected: " + string.Join(", ", selectedValues);
//debugMsg += "\nCalendar Date: " + _CalendarDate;
//debugMsg += _LastDateAvailable != null ? "\nLast Date Logic: " + _LastDateAvailable : "";
//debugMsg += _selectedAverage != null ? "\nRolling Average Column: " + _selectedAverage : "";
//MessageBox.Show(debugMsg);

// Set starting ordinal
int ordinal = cgTable.Name.StartsWith("New Calculation Group") ? 0 :
    (cgTable.CalculationItems.Count > 0 ? cgTable.CalculationItems.Max(ci => ci.Ordinal) + 1 : 0);
        


// ─────────────────────────────────────────────────────────────
// AddCalculationItemIfNotExists helper

List<string> skippedItems = new List<string>();

Action<string, string, int, string, string> AddCalculationItemIfNotExists = 
(name, expression, ordinalPos, description, formatString) =>
{
if (cgTable.CalculationItems.Any(ci => ci.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase)))
{
    skippedItems.Add(name);
    return;
}


    var item = cgTable.AddCalculationItem(name);
    item.Expression = expression;
    item.Ordinal = ordinalPos;
    item.Description = description;
    // FormatDax(item);
};


// ─────────────────────────────────────────────────────────────
// Looping through all selected calculation item categories
foreach (var item in selectedValues)
{

    if (item == "Actual")
    {
        AddCalculationItemIfNotExists("Actual", "SELECTEDMEASURE()", ordinal++, "Actual values", null);

        AddCalculationItemIfNotExists("Actual to date", string.Format(@"
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(),
        _CurrentDates
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Actual value until last date available", null);

        AddCalculationItemIfNotExists("Actual Y-1", 
    string.Format(
        "CALCULATE(\n" +
        "    SELECTEDMEASURE(),\n" +
        "    SAMEPERIODLASTYEAR({0})\n" +
        ")", 
        _CalendarDate
    ), 
    ordinal++, "Actual value last year", null);

        AddCalculationItemIfNotExists("Actual to date Y-1", string.Format(@"
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_CurrentDates)
    )", 
_LastDateAvailable, 
_CalendarDate), ordinal++, "Actual value last year, hiding future dates as of this year", null);

        AddCalculationItemIfNotExists("AOA", string.Format(@"
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _thisyear = 
    CALCULATE(
        SELECTEDMEASURE(),
        _CurrentDates
    )
VAR _lastyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
    
RETURN 
    _thisyear - _lastyear", _LastDateAvailable, _CalendarDate), ordinal++, "AOA = Actual over Actual: Deviation for actual value this year and last year, hiding future dates", null);

        AddCalculationItemIfNotExists("AOA %", string.Format(@"
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _thisyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        _CurrentDates
        )
VAR _lastyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
VAR _deviationYear = _thisyear - _lastyear

RETURN 
    DIVIDE(_deviationYear, _lastyear)", _LastDateAvailable, _CalendarDate), ordinal++, "AOA %: Percentage change this year and last year, hiding future dates", "\"0%\"");

        AddCalculationItemIfNotExists("AOA C", string.Format(@"
VAR _thisyear = SELECTEDMEASURE()
VAR _lastyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR({0})
    )
    
RETURN 
    _thisyear - _lastyear", _CalendarDate), ordinal++, "AOA C = Complete Actual over Actual, not hiding future dates", null);

        AddCalculationItemIfNotExists("AOA C %", string.Format(@"
VAR _thisyear = SELECTEDMEASURE()
VAR _lastyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR({0})
    )
VAR _deviationYear = _thisyear - _lastyear

RETURN 
    DIVIDE(_deviationYear, _lastyear)", _CalendarDate), ordinal++, "AOA C %: Complete Actual over Actual % change, not hiding future dates", "\"0%\"");

        continue;
    }

        // MTD logic
    if (item == "MTD")
    {
        AddCalculationItemIfNotExists("MTD", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(_CurrentDates)
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated Month to date, hiding future dates after last available date", null);

        AddCalculationItemIfNotExists("MTD LY", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated Month to date last year, hiding future dates from last available date (same date last year)", null);

        AddCalculationItemIfNotExists("MOMTD", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentMonth = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(_CurrentDates)
    )
VAR _PreviousMonth = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )
    
RETURN 
    _CurrentMonth - _PreviousMonth", _LastDateAvailable, _CalendarDate), ordinal++, "MOMTD = Month over Month to date: Month-to-Date Deviation, this year and last year accumulated", null);

        AddCalculationItemIfNotExists("MOMTD %", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentMonth = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(_CurrentDates)
    )
VAR _PreviousMonth = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )
VAR _DeltaMonth = _CurrentMonth - _PreviousMonth

RETURN 
    DIVIDE(_DeltaMonth, _PreviousMonth)", _LastDateAvailable, _CalendarDate), ordinal++, "MOMTD = Month over Month to date %: % change this year compared to last year accumulated", "\"0%\"");

        AddCalculationItemIfNotExists("MTD PM", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(PREVIOUSMONTH(_CurrentDates))
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated Month to date previous Month, hiding future dates from last available date (same date last Month)", null);

        AddCalculationItemIfNotExists("MTD PQ", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(PREVIOUSQUARTER(_CurrentDates))
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated Month to date previous Quarter, hiding future dates from last available date (same date last quarter)", null);

        AddCalculationItemIfNotExists("MTD C", string.Format(@"CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD({0})
        )"
        , _CalendarDate), ordinal++, "MTD C = Month to date Complete: accumulated, without hiding of future dates", null);

        AddCalculationItemIfNotExists("MTD LY C", string.Format(@"CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(SAMEPERIODLASTYEAR({0}))
        )", 
        _CalendarDate), ordinal++, "MTD LY C = Month to date Complete Last Year: Month last year accumulated, without hiding future dates", null);

        AddCalculationItemIfNotExists("MTD PM C", string.Format(@"CALCULATE(
        SELECTEDMEASURE(), 
        DATESMTD(PREVIOUSMONTH({0}))
        )", _CalendarDate), ordinal++, "MTD PM C = Month to date, Previous month, Complete: Last Month previous month accumulated, without hiding future dates", null);

        continue;
    }
    
 // QTD logic
    if (item == "QTD")
    {
        AddCalculationItemIfNotExists("QTD", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
    
RETURN 
    
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(_CurrentDates)
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated quarter to date, hiding future dates after last available date", null);

        AddCalculationItemIfNotExists("QTD LY", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
        
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )

RETURN 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )", _LastDateAvailable, _CalendarDate), ordinal++, "Accumulated quarter to date last year, hiding future dates from last available date (same date last year)", null);

        AddCalculationItemIfNotExists("QOQTD", string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
      
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentYear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(_CurrentDates)
     )
VAR _PreviousYear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )

RETURN 
    _CurrentYear - _PreviousYear", _LastDateAvailable, _CalendarDate), ordinal++, "QOQTD = Quarter over Quarter to date: Quarter-to-Date Deviation, this year and last year accumulated", null);

        AddCalculationItemIfNotExists("QOQTD %", string.Format(@"
// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentYear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(_CurrentDates)
    )
VAR _PreviousYear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESQTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear

RETURN 
    DIVIDE(_DeltaYear, _PreviousYear)", _LastDateAvailable, _CalendarDate), ordinal++, "QOQTD = Quarter over Quarter to date: Quarter-to-Date Index, % change this year compared to last year accumulated", "\"0%\"");

        AddCalculationItemIfNotExists("QTD C", string.Format(@"CALCULATE(
     SELECTEDMEASURE(), 
     DATESQTD({0})
     )", _CalendarDate), ordinal++, "QTD C = Quarter to date Complete: Quarter to date accumulated, without hiding future dates", null);

        AddCalculationItemIfNotExists("QTD LY C", string.Format(@"CALCULATE(
    SELECTEDMEASURE(), 
    DATESQTD(SAMEPERIODLASTYEAR({0}))
    )", _CalendarDate), ordinal++, "QTD LY C = Quarter to date last year Complete: Last Quarter accumulated, without hiding future dates", null);

        continue;
    }

// YTD logic
if (item == "YTD")
{
    AddCalculationItemIfNotExists(
        "YTD",
        string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0}
VAR _CurrentDates = 
    FILTER(VALUES({1}),
    {1} <= _LastDayAvailable
    )
VAR _Result = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESYTD(_CurrentDates)
    )
    
RETURN 
    _Result", _LastDateAvailable, _CalendarDate),
        ordinal++, "Accumulated year to date, hiding future dates after last available date", null
    );

    AddCalculationItemIfNotExists(
        "LYTD",
        string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0}
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _Result = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESYTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )
    
RETURN 
    _Result", _LastDateAvailable, _CalendarDate),
        ordinal++, "Accumulated year to date last year, hiding future dates from last available date (same date last year)", null
    );

    AddCalculationItemIfNotExists(
        "YOYTD",
        string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0}
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentYear = 
     CALCULATE(
        SELECTEDMEASURE(), 
        DATESYTD(_CurrentDates)
     )
VAR _PreviousYear = 
     CALCULATE(
        SELECTEDMEASURE(), 
        DATESYTD(SAMEPERIODLASTYEAR(_CurrentDates))
     )
     
RETURN 
    _CurrentYear - _PreviousYear", _LastDateAvailable, _CalendarDate),
        ordinal++, "YOYTD = Year-to-Date Deviation, this year and last year accumulated", null
    );

    AddCalculationItemIfNotExists(
        "YOYTD %",
        string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/

VAR _LastDayAvailable = {0}
VAR _CurrentDates = 
    FILTER(VALUES({1}), 
    {1} <= _LastDayAvailable
    )
VAR _CurrentYear = 
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(_CurrentDates)
    )
VAR _PreviousYear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        DATESYTD(SAMEPERIODLASTYEAR(_CurrentDates))
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear

RETURN 
    DIVIDE(_DeltaYear, _PreviousYear)", _LastDateAvailable, _CalendarDate),
        ordinal++, "YOYTD = Year-to-Date Index %, % change this year vs last", "\"0%\""
    );

    AddCalculationItemIfNotExists(
        "YTD C",
        string.Format(@"CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD({0})
        )", _CalendarDate),
        ordinal++, "YTD C = Year-to-Date Complete, no future-date filtering", null
    );

    AddCalculationItemIfNotExists(
        "LYTD C",
        string.Format(@"CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(SAMEPERIODLASTYEAR({0}))
        )", _CalendarDate),
        ordinal++, "LYTD C = Last Year-to-Date Complete, no future-date filtering", null
    );

    continue;
}

      if (item == "R12")
    {
        AddCalculationItemIfNotExists("Rolling 12M", string.Format(@"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/

VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = 
    CALCULATE(
        MAX({1}),
        FILTER(VALUES({1}), 
        {1} <= _LastDayShown
        )
    )
VAR _PreviousDates =
    DATESINPERIOD({1}, _ReferenceDate, -_NumOfMonths, MONTH)

VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(), 
        _PreviousDates
    )
VAR _firstDateInPeriod = MINX(_PreviousDates, {1})
RETURN IF(_firstDateInPeriod <= _ReferenceDate, _Result)", _LastDateAvailable, _CalendarDate), ordinal++, "Rolling 12 months, hiding future dates after last available date", null);
        
AddCalculationItemIfNotExists("Rolling 12M LY", string.Format(@"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12

VAR _ReferenceDate = 
    CALCULATE(
        MAX({1}),
        FILTER(VALUES({1}), 
        {1} <= _LastDayShown
        )
    )
    
VAR _PreviousDates =
    DATESINPERIOD({1}, _ReferenceDate, -_NumOfMonths, MONTH)
    
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _firstDateInPeriod = MINX(_PreviousDates, {1})

RETURN 
    IF(_firstDateInPeriod <= _ReferenceDate, _Result)", _LastDateAvailable, _CalendarDate), ordinal++, "Rolling 12 months last year, hiding future dates", null);
       
AddCalculationItemIfNotExists("Rolling 12M Dev", string.Format(@"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = 
    CALCULATE(
        MAX({1}),
        FILTER(VALUES({1}), 
        {1} <= _LastDayShown
        )
    )
    
VAR _PreviousDates =
    DATESINPERIOD({1}, _ReferenceDate, -_NumOfMonths, MONTH)
    
VAR _CurrentResult =
    CALCULATE(
        SELECTEDMEASURE(), 
        _PreviousDates)
        
VAR _PreviousResult =
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
    
VAR _firstDateInPeriod = MINX(_PreviousDates, {1})

RETURN 
    IF(_firstDateInPeriod <= _ReferenceDate, _CurrentResult - _PreviousResult)", _LastDateAvailable, _CalendarDate), ordinal++, "Rolling 12 months Deviation", null);
        
AddCalculationItemIfNotExists("Rolling 12M idx", string.Format(@"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = 
    CALCULATE(
        MAX({1}),
        FILTER(VALUES({1}), 
        {1} <= _LastDayShown
        )
    )
    
VAR _PreviousDates =
    DATESINPERIOD({1}, _ReferenceDate, -_NumOfMonths, MONTH)

VAR _CurrentResult =
    CALCULATE(
        SELECTEDMEASURE(), 
        _PreviousDates
    )
VAR _PreviousResult =
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
    
VAR _DeviationResult = _CurrentResult - _PreviousResult
VAR _firstDateInPeriod = MINX(_PreviousDates, {1})

RETURN 
    IF(_firstDateInPeriod <= _ReferenceDate, DIVIDE(_DeviationResult, _PreviousResult))", _LastDateAvailable, _CalendarDate), ordinal++, "Rolling 12 months Index %", "\"0%\"");
        continue;
    }
    
    
    
    if (item == "Rolling average")
    {
        if (!string.IsNullOrEmpty(_selectedAverage))
        {
            AddCalculationItemIfNotExists("Rolling 12M avg", string.Format(@"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/

 //Calculating the average based on the selection in VALUES, such as per day or per month
 
VAR _LastDayShown = {0}
        
VAR _NumOfMonths = 12
// Hiding future days. If not remove the calculate in _ReferenceDate
VAR  _LastCurrentDate = CALCULATE(
    max({1}), FILTER(
        VALUES( {1} ),
        {1} <= _LastDayShown
    )
)

 VAR _Period =
        DATESINPERIOD ( {1}, _LastCurrentDate, - _NumOfMonths, MONTH )  
    
VAR _Result =
        CALCULATE (
            AVERAGEX (
                VALUES ( {2} ),     
                SELECTEDMEASURE () 
            ),
            _Period
        ) 
    
     VAR _firstDateInPeriod = MINX ( _Period, {1} )
   
    
RETURN 
    IF ( _firstDateInPeriod <= _LastCurrentDate, _Result )",
    _LastDateAvailable, _CalendarDate, _selectedAverage),
    ordinal++,
    string.Format("Rolling 12 months average per {0}, hiding future dates", _selectedAverage),
    null);
        }
        continue;
    }
    
        if (item == "Rolling total")
        {     
    AddCalculationItemIfNotExists("Rolling Total", string.Format(@"
var _currdate= {0}

RETURN
    CALCULATE(
        SELECTEDMEASURE(),
        FILTER(
            ALLSELECTED({1}),
            ISONORAFTER({1}, _currdate, DESC)
        )
    )

     "
        , _LastDateAvailable, _CalendarDate), ordinal++,
        "Running Total, from the first date available until selected date", null);
        continue;
    }
    
    

    
    // Placeholder: additional Time Intelligence logic goes here
  //  AddCalculationItemIfNotExists(item, "SELECTEDMEASURE()", ordinal++, "Time Intelligence - " + item, null);
}

if (skippedItems.Count > 0)
{
string message = "The following calculation item(s) were skipped because they already exist:\n\n" + string.Join("\n", skippedItems);

    MessageBox.Show(message, "Items Skipped", MessageBoxButtons.OK, MessageBoxIcon.Information);
}


