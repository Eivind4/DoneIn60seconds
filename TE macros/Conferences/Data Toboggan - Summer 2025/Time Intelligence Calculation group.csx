// This script includes the select time calculations that are to be included. The idea is from the following script, but here it is included as a Calculation Group instead and with some modifications
//        Data-Goblin inspiration: https://github.com/data-goblin/powerbi-macguyver-toolbox/blob/main/tabular-editor-scripts/csharp-scripts/add-dax-templates/add-time-intelligence.csx

// Context for macro:
//        - Model
//        - Calculation Group Table


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
        Height = 220, // Increased from 150 to 220
        FormBorderStyle = FormBorderStyle.FixedDialog,
        Text = caption,
        StartPosition = FormStartPosition.CenterScreen
    };
    
    Label textLabel = new Label() 
    { 
        Left = 50, 
        Top = 20, 
        Width = 500, // Increased width to match form
        Height = 80, // Increased height for multi-line text
        Text = text, 
        AutoSize = false // Set to false to allow custom sizing
    };
    
    TextBox textBox = new TextBox() 
    { 
        Left = 50, 
        Top = 110, // Moved down to accommodate larger label
        Width = 500, 
        Text = defaultValue 
    };
    
    Button confirmation = new Button() 
    { 
        Text = "OK", 
        Left = 450, 
        Width = 100, 
        Top = 140, // Moved down accordingly
        DialogResult = DialogResult.OK 
    };
    
    confirmation.Click += (sender, e) => { prompt.Close(); };
    prompt.Controls.Add(textLabel);
    prompt.Controls.Add(textBox);
    prompt.Controls.Add(confirmation);
    prompt.AcceptButton = confirmation;

    return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
};

// Get the selected calculation group table. Checks the name if it is a new or existing calculation group
CalculationGroupTable cgTable = null;

// Check if a calculation group table is selected
if (Selected.Tables.Any() && Selected.Tables.First() is CalculationGroupTable)
{
    cgTable = Selected.Tables.First() as CalculationGroupTable;
}
else
{
    // No calculation group selected, check if any exist in the model
    var existingCalcGroups = Model.Tables.OfType<CalculationGroupTable>().ToList();
    
    if (existingCalcGroups.Any())
    {
        // Show dialog to select existing calculation group or create new one
        Form selectForm = new Form();
        selectForm.Text = "Select Calculation Group";
        selectForm.Width = 450;
        selectForm.Height = 300;
        selectForm.FormBorderStyle = FormBorderStyle.FixedDialog;
        selectForm.StartPosition = FormStartPosition.CenterScreen;

        Label instructionLabel = new Label()
        {
            Left = 20,
            Top = 20,
            Width = 400,
            Height = 40,
            Text = "No calculation group table selected. Choose an option:"
        };

        Button selectExistingButton = new Button()
        {
            Text = "Select Existing",
            Left = 50,
            Width = 120,
            Height = 35,
            Top = 80,
            DialogResult = DialogResult.Yes
        };

        Button createNewButton = new Button()
        {
            Text = "Create New",
            Left = 180,
            Width = 120,
            Height = 35,
            Top = 80,
            DialogResult = DialogResult.No
        };

        Button cancelButton = new Button()
        {
            Text = "Cancel",
            Left = 310,
            Width = 80,
            Height = 35,
            Top = 80,
            DialogResult = DialogResult.Cancel
        };

        selectForm.Controls.Add(instructionLabel);
        selectForm.Controls.Add(selectExistingButton);
        selectForm.Controls.Add(createNewButton);
        selectForm.Controls.Add(cancelButton);

        DialogResult choice = selectForm.ShowDialog();

        if (choice == DialogResult.Cancel)
        {
            return; // Exit script
        }
        else if (choice == DialogResult.Yes)
        {
            // Select existing calculation group
            Form existingForm = new Form();
            ListBox cgListBox = new ListBox();
            Button okButton = new Button();
            Button cancelBtn = new Button();

            existingForm.Text = "Select Existing Calculation Group";
            existingForm.Width = 400;
            existingForm.Height = 300;
            existingForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            existingForm.StartPosition = FormStartPosition.CenterScreen;

            cgListBox.Width = 360;
            cgListBox.Height = 180;
            cgListBox.Left = 10;
            cgListBox.Top = 20;
            cgListBox.SelectionMode = SelectionMode.One;
            cgListBox.DataSource = existingCalcGroups.Select(cg => cg.Name).ToList();

            okButton.Text = "OK";
            okButton.Left = 220;
            okButton.Top = 220;
            okButton.DialogResult = DialogResult.OK;

            cancelBtn.Text = "Cancel";
            cancelBtn.Left = 300;
            cancelBtn.Top = 220;
            cancelBtn.DialogResult = DialogResult.Cancel;

            existingForm.Controls.Add(cgListBox);
            existingForm.Controls.Add(okButton);
            existingForm.Controls.Add(cancelBtn);

            if (existingForm.ShowDialog() == DialogResult.OK && cgListBox.SelectedIndex >= 0)
            {
                string selectedName = cgListBox.SelectedItem.ToString();
                cgTable = existingCalcGroups.FirstOrDefault(cg => cg.Name == selectedName);
            }
            else
            {
                return; // User cancelled
            }
        }
        else
        {
            // Create new calculation group
            cgTable = Model.AddCalculationGroup("New Calculation Group");
        }
    }
    else
    {
        // No calculation groups exist, create new one
        cgTable = Model.AddCalculationGroup("New Calculation Group");
    }
}

// Verify we have a calculation group table
if (cgTable == null)
{
    Error("No calculation group table available. Please select or create a calculation group table first.");
    return;
}

// Rename Calculation Group Table and Column if it's a new one
if (cgTable.Name.StartsWith("New Calculation Group"))
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

    form.Text = "Select Column";
    form.Width = 400;
    form.Height = 470;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.StartPosition = FormStartPosition.CenterScreen;

    label.Text = title;
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

    form.Text = "Select Measure";
    form.Width = 400;
    form.Height = 470;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.StartPosition = FormStartPosition.CenterScreen;

    label.Text = title;
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
    "Actual", "MTD", "QTD", "YTD", "R12", "Rolling average", "Rolling total", "Start and End"
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

// ─────────────────────────────────────────────────────────────
// NEW: Future Date Filter Logic Selection

string _LastDateAvailable = null;
string futureDateFilterType = "Measure"; // Default

bool requiresFutureDateLogic = selectedValues.Any(x => new[] { "Actual", "MTD", "QTD", "YTD", "R12", "Rolling average", "Rolling total" }.Contains(x));

if (requiresFutureDateLogic)
{
    // Show dialog to choose between Column and Measure approach
    Form futureLogicForm = new Form();
    futureLogicForm.Text = "Hide Future Dates";
    futureLogicForm.Width = 500;
    futureLogicForm.Height = 300;
    futureLogicForm.FormBorderStyle = FormBorderStyle.FixedDialog;
    futureLogicForm.StartPosition = FormStartPosition.CenterScreen;

    Label promptLabel = new Label()
    {
        Left = 20,
        Top = 20,
        Width = 450,
        Height = 20,
        Text = "Choose method for hiding future dates:"
    };

    Label columnInfoLabel = new Label()
    {
        Left = 20,
        Top = 45,
        Width = 450,
        Height = 40,
        Text = "• Column: Requires column in date-table to filter dates (i.e. Is Historical (True/False). Likely to be more performant",
        AutoSize = false
    };

    Label measureInfoLabel = new Label()
    {
        Left = 20,
        Top = 85,
        Width = 450,
        Height = 40,
        Text = "• Measure: Select a measure (Measure preferred, but not required). More flexible if no column to filter dates",
        AutoSize = false
    };

    Button columnButton = new Button()
    {
        Text = "Column",
        Left = 70,
        Width = 120,
        Height = 35,
        Top = 140,
        DialogResult = DialogResult.Yes
    };

    Button measureButton = new Button()
    {
        Text = "Measure",
        Left = 200,
        Width = 120,
        Height = 35,
        Top = 140,
        DialogResult = DialogResult.No
    };

    Button cancelButton = new Button()
    {
        Text = "Cancel",
        Left = 330,
        Width = 120,
        Height = 35,
        Top = 140,
        DialogResult = DialogResult.Cancel
    };

    futureLogicForm.Controls.Add(promptLabel);
    futureLogicForm.Controls.Add(columnInfoLabel);
    futureLogicForm.Controls.Add(measureInfoLabel);
    futureLogicForm.Controls.Add(columnButton);
    futureLogicForm.Controls.Add(measureButton);
    futureLogicForm.Controls.Add(cancelButton);

    futureLogicForm.AcceptButton = columnButton;
    futureLogicForm.CancelButton = cancelButton;

    DialogResult futureResult = futureLogicForm.ShowDialog();

    if (futureResult == DialogResult.Cancel)
    {
        return;
    }

    futureDateFilterType = (futureResult == DialogResult.Yes) ? "Column" : "Measure";

    if (futureDateFilterType == "Column")
    {
        // Get columns from the date table
        var dateTable = Model.Tables[_calendarTableName];
        var dateTableColumns = dateTable.Columns.ToList();
        
        var selectedColumn = SelectColumnFromList(dateTableColumns, "Select the COLUMN to filter future dates\n(e.g., Is Historical)");
        
        if (selectedColumn == null)
        {
            Error("No column selected for filtering future dates.");
            return;
        }
        
        // Ask user for the column value to filter by - with larger form
        Form prompt = new Form()
        {
            Width = 600,
            Height = 220, // Increased height
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = "Column Filter Value",
            StartPosition = FormStartPosition.CenterScreen
        };
        
        Label textLabel = new Label() 
        { 
            Left = 50, 
            Top = 20, 
            Width = 500,
            Height = 80,
            Text = "Enter the VALUE that the column should equal to show historical dates:\n\nExamples:\n• TRUE (for boolean columns)\n• 1 (for integer columns)\n• 'Yes' (for text columns - include quotes)",
            AutoSize = false
        };
        
        TextBox textBox = new TextBox() 
        { 
            Left = 50, 
            Top = 110,
            Width = 500, 
            Text = "TRUE" 
        };
        
        Button confirmation = new Button() 
        { 
            Text = "OK", 
            Left = 450, 
            Width = 100, 
            Top = 140,
            DialogResult = DialogResult.OK 
        };
        
        confirmation.Click += (sender, e) => { prompt.Close(); };
        prompt.Controls.Add(textLabel);
        prompt.Controls.Add(textBox);
        prompt.Controls.Add(confirmation);
        prompt.AcceptButton = confirmation;

        string columnValue = prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        
        if (string.IsNullOrEmpty(columnValue))
        {
            Error("No column value provided for filtering future dates.");
            return;
        }
        
        _LastDateAvailable = selectedColumn.DaxObjectFullName + " = " + columnValue;
    }
    else
    {
        // Original measure-based logic
        var _dateFormattedMeasures = Model.AllMeasures
            .Where(m => m.FormatString != null && (m.FormatString.Contains("yy") || m.FormatString.Contains("yyyy")))
            .ToList();

        var selectedMeasure = SelectMeasureFromList(_dateFormattedMeasures, "Select MEASURE for hiding future dates\n\n(Cancel will create logic for last date)");
        _LastDateAvailable = selectedMeasure != null
            ? selectedMeasure.DaxObjectName
            : string.Format("MIN(MAX({0}), TODAY())", _CalendarDate);
    }
}

string _selectedAverage = null;
if (selectedValues.Contains("Rolling average"))
{
    var _selectedTableName = _CalendarDate.Split('[')[0].Trim('\'');
    var _selectedDateTable = Model.Tables[_selectedTableName];
    var _columnsInDateTable = _selectedDateTable.Columns.ToList();

    var avgCol = SelectColumnFromList(_columnsInDateTable, "Select column for Rolling 12M Average\n\n(e.g., daily/monthly)");
    if (avgCol != null)
    {
        _selectedAverage = avgCol.DaxObjectFullName;
    }
}

string _ISO_year = null;
string _ISO_week = null;
string _dayNoOfWeek = null;
string _ISO_yearWeekRunningNumber = null;


string _FromDate = null;
string _EndDate = null;

if (selectedValues.Contains("Start and End"))
{
    // Creates a list of columns of type date that is NOT marked as key (not the main date table column)
    var _dateColumnsNotCalendar = Model.AllColumns.Where(c => c.DataType == DataType.DateTime &&
        c.IsKey != true).ToList();

    try
    {
        // Select the Start date column
        var fromDateCol = SelectColumnFromList(_dateColumnsNotCalendar, "START AND END DATE CALCULATIONS:\nSelect the START-date column to be used\n(Assumption: DateTime formatted)");
        if (fromDateCol == null)
        {
            Error("No Start date column selected. Start and End calculations will be skipped.");
        }
        else
        {
            _FromDate = fromDateCol.DaxObjectFullName;
            
            // Select the End date column
            var endDateCol = SelectColumnFromList(_dateColumnsNotCalendar, "START AND END DATE CALCULATIONS:\nSelect the END-date column to be used\n(Assumption: DateTime formatted)");
            if (endDateCol == null)
            {
                Error("No End date column selected. Start and End calculations will be skipped.");
                _FromDate = null; // Reset since we need both
            }
            else
            {
                _EndDate = endDateCol.DaxObjectFullName;
            }
        }
    }
    catch
    {
        Error("Error selecting Start and End date columns. Start and End calculations will be skipped.");
        _FromDate = null;
        _EndDate = null;
    }
}


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
    item.FormatStringExpression = formatString;
};

// ─────────────────────────────────────────────────────────────
// NEW: Helper function for time intelligence items with future date hiding

Action<string, string, string, int, string, string> AddTimeIntelligenceItemWithHideFuture = 
(name, daxColumnTemplate, daxMeasureTemplate, ordinalPos, description, formatString) =>
{
    string expression = futureDateFilterType == "Column"
        ? string.Format(daxColumnTemplate, _LastDateAvailable, _CalendarDate)
        : string.Format(daxMeasureTemplate, _LastDateAvailable, _CalendarDate);

    // Make description dynamic based on filter type
    string dynamicDescription = futureDateFilterType == "Column"
        ? description.Replace("{FILTER_DESCRIPTION}", string.Format("after column: {0}", _LastDateAvailable))
        : description.Replace("{FILTER_DESCRIPTION}", string.Format("after last date returned from measure: {0}", _LastDateAvailable));

    AddCalculationItemIfNotExists(name, expression, ordinalPos, dynamicDescription, formatString);
};

// Helper function that passes the third parameter for rolling average
Action<string, string, string, int, string, string> AddRollingAverageItem = 
(name, daxColumnTemplate, daxMeasureTemplate, ordinalPos, description, formatString) =>
{
    string expression = futureDateFilterType == "Column"
        ? string.Format(daxColumnTemplate, _LastDateAvailable, _CalendarDate, _selectedAverage)
        : string.Format(daxMeasureTemplate, _LastDateAvailable, _CalendarDate, _selectedAverage);

    // Make description dynamic based on filter type
    string dynamicDescription = futureDateFilterType == "Column"
        ? description.Replace("{FILTER_DESCRIPTION}", string.Format("after column: {0}", _LastDateAvailable))
        : description.Replace("{FILTER_DESCRIPTION}", string.Format("after last date returned from measure: {0}", _LastDateAvailable));

    AddCalculationItemIfNotExists(name, expression, ordinalPos, dynamicDescription, formatString);
};

// ─────────────────────────────────────────────────────────────
// Looping through all selected calculation item categories
foreach (var item in selectedValues)
{
    if (item == "Actual")
    {
        AddCalculationItemIfNotExists("Actual", "SELECTEDMEASURE()", ordinal++, "Actual values", null);

        AddTimeIntelligenceItemWithHideFuture("Actual to date", 
            @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
RETURN _Result",

            @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
RETURN _Result",
            ordinal++, "Actual value until last date available, hiding future dates {FILTER_DESCRIPTION}", null);

        AddCalculationItemIfNotExists("Actual Y-1", 
            string.Format(
                "CALCULATE(\n" +
                "    SELECTEDMEASURE(),\n" +
                "    SAMEPERIODLASTYEAR({0})\n" +
                ")", 
                _CalendarDate
            ), 
            ordinal++, "Actual value last year", null);

        AddTimeIntelligenceItemWithHideFuture("Actual to date Y-1",
            @"VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
RETURN _Result",

            @"VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
RETURN _Result",
            ordinal++, "Actual value last year, hiding future dates {FILTER_DESCRIPTION}", null);

        AddTimeIntelligenceItemWithHideFuture("AOA",
            @"VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _thisyear = 
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
RETURN _thisyear - _lastyear",

            @"VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _thisyear = 
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
RETURN _thisyear - _lastyear",
            ordinal++, "AOA = Actual over Actual: Deviation for actual value this year and last year, hiding future dates {FILTER_DESCRIPTION}", null);

        AddTimeIntelligenceItemWithHideFuture("AOA %",
            @"VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _thisyear = 
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
VAR _deviationYear = _thisyear - _lastyear
RETURN DIVIDE(_deviationYear, _lastyear)",

            @"VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _thisyear = 
    CALCULATE (
        SELECTEDMEASURE(),
        _CurrentDates
    )
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_CurrentDates)
    )
VAR _deviationYear = _thisyear - _lastyear
RETURN DIVIDE(_deviationYear, _lastyear)",
            ordinal++, "AOA %: Percentage change this year and last year, hiding future dates {FILTER_DESCRIPTION}", "\"0%\"");

        AddCalculationItemIfNotExists("AOA C", string.Format(@"VAR _thisyear = SELECTEDMEASURE()
VAR _lastyear = 
    CALCULATE(
        SELECTEDMEASURE(), 
        SAMEPERIODLASTYEAR({0})
    )
    
RETURN 
    _thisyear - _lastyear", _CalendarDate), ordinal++, "AOA C = Complete Actual over Actual, not hiding future dates", null);

        AddCalculationItemIfNotExists("AOA C %", string.Format(@"VAR _thisyear = SELECTEDMEASURE()
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
     AddTimeIntelligenceItemWithHideFuture(
        "MTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
RETURN _Result",
        ordinal++,
        "Accumulated Month to date, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "MTD LY",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",
        ordinal++,
        "Accumulated Month to date last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "MOMTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentMonth =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
VAR _PreviousMonth =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentMonth - _PreviousMonth",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentMonth =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
VAR _PreviousMonth =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentMonth - _PreviousMonth",
        ordinal++,
        "MOMTD = Month over Month to date: Month-to-Date Deviation, this year and last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "MOMTD %",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentMonth =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
VAR _PreviousMonth =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaMonth = _CurrentMonth - _PreviousMonth
RETURN DIVIDE(_DeltaMonth, _PreviousMonth)",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentMonth =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESMTD (_CurrentDates)
    )
VAR _PreviousMonth =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaMonth = _CurrentMonth - _PreviousMonth
RETURN DIVIDE(_DeltaMonth, _PreviousMonth)",
        ordinal++,
        "MOMTD = Month over Month to date %: % change this year compared to last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        "\"0%\""
    );

   AddTimeIntelligenceItemWithHideFuture(
        "MTD PM",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            DATEADD( _CurrentDates,-1,MONTH )
        )
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            DATEADD( _CurrentDates,-1,MONTH )
        )
    )
RETURN _Result",
        ordinal++,
        "Accumulated Month to date previous Month, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

   AddTimeIntelligenceItemWithHideFuture(
        "MTD PQ",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            DATEADD( _CurrentDates,-3,MONTH )
        )
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESMTD(
            DATEADD( _CurrentDates,-3,MONTH )
        )
    )
RETURN _Result",
        ordinal++,
        "Accumulated Month to date previous Quarter, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddCalculationItemIfNotExists(
        "MTD C",
        string.Format(@"CALCULATE (
    SELECTEDMEASURE (),
    DATESMTD ({0})
)", _CalendarDate),
        ordinal++,
        "MTD C = Month to date Complete: accumulated, without hiding of future dates",
        null
    );

    AddCalculationItemIfNotExists(
        "MTD LY C",
        string.Format(@"CALCULATE(
    SELECTEDMEASURE(),
    DATESMTD(
        SAMEPERIODLASTYEAR({0})
    )
)", _CalendarDate),
        ordinal++,
        "MTD C = Month to date Complete: Last Month last year accumulated, without hiding future dates",
        null
    );

    AddCalculationItemIfNotExists(
        "MTD PM C",
        string.Format(@"CALCULATE(
    SELECTEDMEASURE(),
    DATESMTD(
        PREVIOUSMONTH({0})
    )
)", _CalendarDate),
        ordinal++,
        "MTD PM C = Month to date, Previous month, Complete: Last Month previous month accumulated, without hiding future dates",
        null
    );

        }   //END - MTD

    //QTD logic
    if (item == "QTD")
      {
    AddTimeIntelligenceItemWithHideFuture(
        "QTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
RETURN _Result",
        ordinal++,
        "Accumulated quarter to date, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "QTD LY",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",
        ordinal++,
        "Accumulated quarter to date last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "QOQTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentYear - _PreviousYear",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentYear - _PreviousYear",
        ordinal++,
        "QOQTD = Quarter over Quarter to date: Quarter-to-Date Deviation, this year and last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "QOQTD %",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESQTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESQTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)",
        ordinal++,
        "QOQTD = Quarter over Quarter to date: Quarter-to-Date Index, % change this year compared to last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        "\"0%\""
    );

    AddCalculationItemIfNotExists(
        "QTD C",
        string.Format(@"CALCULATE (
    SELECTEDMEASURE (),
    DATESQTD ({0})
)", _CalendarDate),
        ordinal++,
        "QTD C= Quarter to date Complete: Quarter to date accumulated, without hiding future dates",
        null
    );

    AddCalculationItemIfNotExists(
        "QTD LY C",
        string.Format(@"CALCULATE(
    SELECTEDMEASURE(),
    DATESQTD(
        SAMEPERIODLASTYEAR({0})
    )
)", _CalendarDate),
        ordinal++,
        "QTD LY C= Quarter to date last year Complete:Last Quarter accumulated, without hiding future dates",
        null
    );

    }  
    
    //YTD logic
    if (item == "YTD")
    {
      AddTimeIntelligenceItemWithHideFuture(
        "YTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
RETURN _Result",
        ordinal++,
        "Accumulated year to date, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "LYTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _Result",
        ordinal++,
        "Accumulated year to date last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "YOYTD",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentYear - _PreviousYear",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
RETURN _CurrentYear - _PreviousYear",
        ordinal++,
        "YOYTD = Year over year to date: Year-to-Date Deviation, this year and last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "YOYTD %",
        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)",

        @"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {0} 
VAR _CurrentDates = 
    FILTER(
        VALUES({1}),
        {1} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates)
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates)
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)",
        ordinal++,
        "YOYTD = Year over year to date: Year-to-Date Index, % change this year compared to last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
        "\"0%\""
    );

    AddCalculationItemIfNotExists(
        "YTD C",
        string.Format(@"CALCULATE (
    SELECTEDMEASURE (),
    DATESYTD ({0})
)", _CalendarDate),
        ordinal++,
        "YTD C = Year to date Complete: Year to date accumulated, without hiding future dates",
        null
    );

    AddCalculationItemIfNotExists(
        "LYTD C",
        string.Format(@"CALCULATE(
    SELECTEDMEASURE(),
    DATESYTD(
        SAMEPERIODLASTYEAR({0})
    )
)", _CalendarDate),
        ordinal++,
        "LYTD C = Last year to date Complete: Last Year accumulated, without hiding future dates",
        null
    );

    }  //END YTD

    //R12 logic
    if (item == "R12")

    {
 AddTimeIntelligenceItemWithHideFuture(
        "Rolling 12M",
        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}),
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _Result )",

        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}), FILTER(
        VALUES( {1} ),
        {1} <= _LastDayShown
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _Result )",
        ordinal++,
        "Rolling 12 months, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "Rolling 12M LY",
        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}),
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _Result )",

        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}), FILTER(
        VALUES( {1} ),
        {1} <= _LastDayShown
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _Result )",
        ordinal++,
        "Rolling 12 months last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "Rolling 12M Dev",
        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}),
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _CurrentResult =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _PreviousResult =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _CurrentResult - _PreviousResult )",

        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}), FILTER(
        VALUES( {1} ),
        {1} <= _LastDayShown
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _CurrentResult =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _PreviousResult =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, _CurrentResult - _PreviousResult )",
        ordinal++,
        "Rolling 12 months Deviation, this year and last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "Rolling 12M idx",
        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}),
    CALCULATETABLE(
        VALUES({1}),
        {0}
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _CurrentResult =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _PreviousResult =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _DeviationResult = _CurrentResult - _PreviousResult
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, DIVIDE( _DeviationResult, _PreviousResult ) )",

        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _ReferenceDate = CALCULATE(
    MAX({1}), FILTER(
        VALUES( {1} ),
        {1} <= _LastDayShown
    )
)
VAR _PreviousDates =
    DATESINPERIOD (
        {1},
        _ReferenceDate,
        -_NumOfMonths,
        MONTH
    )
VAR _CurrentResult =
    CALCULATE (
        SELECTEDMEASURE(),
        _PreviousDates
    )
VAR _PreviousResult =
    CALCULATE (
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR(_PreviousDates)
    )
VAR _DeviationResult = _CurrentResult - _PreviousResult
VAR _firstDateInPeriod = MINX ( _PreviousDates, {1} )
RETURN 
    IF ( _firstDateInPeriod <= _ReferenceDate, DIVIDE( _DeviationResult, _PreviousResult ) )",
        ordinal++,
        "Rolling 12 months, % change this year and last year (idx), hiding future dates {FILTER_DESCRIPTION}",
        "\"0%\""
    );

    }
    
        //R12 average logic
   if (item == "Rolling average")
{
    if (!string.IsNullOrEmpty(_selectedAverage))
    {
        AddRollingAverageItem("Rolling 12M avg", 
            @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
// Calculating the average based on the selection in VALUES, such as per day or per month
// Column-based future date filtering

VAR _NumOfMonths = 12
VAR _LastCurrentDate = 
    CALCULATE(
        MAX({1}), 
        CALCULATETABLE(
            VALUES({1}),
            {0}
        )
    )

VAR _Period =
    DATESINPERIOD({1}, _LastCurrentDate, -_NumOfMonths, MONTH)  
    
VAR _Result =
    CALCULATE(
        AVERAGEX(
            VALUES({2}),     
            SELECTEDMEASURE()
        ),
        _Period
    ) 
    
VAR _firstDateInPeriod = MINX(_Period, {1})
   
RETURN 
    IF(_firstDateInPeriod <= _LastCurrentDate, _Result)",

            @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
// Calculating the average based on the selection in VALUES, such as per day or per month
// Measure-based future date filtering
 
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _LastCurrentDate = 
    CALCULATE(
        MAX({1}), 
        FILTER(
            VALUES({1}),
            {1} <= _LastDayShown
        )
    )

VAR _Period =
    DATESINPERIOD({1}, _LastCurrentDate, -_NumOfMonths, MONTH)  
    
VAR _Result =
    CALCULATE(
        AVERAGEX(
            VALUES({2}),     
            SELECTEDMEASURE()
        ),
        _Period
    ) 
    
VAR _firstDateInPeriod = MINX(_Period, {1})
   
RETURN 
    IF(_firstDateInPeriod <= _LastCurrentDate, _Result)",
            ordinal++,
            string.Format("Rolling 12 months average per {0}, hiding future dates {{FILTER_DESCRIPTION}}", _selectedAverage),
            null);
    }
    continue;
} 
        
        //Rolling total
        if (item == "Rolling total")
               {
        AddTimeIntelligenceItemWithHideFuture(
        "Running Total",
        @"
VAR _currdate =
    CALCULATE(
        MAX({1}),
        CALCULATETABLE(
            VALUES({1}),
            {0}
        )
    )
RETURN
    CALCULATE(
        SELECTEDMEASURE(),
        FILTER(
            ALLSELECTED({1}),
            ISONORAFTER({1}, _currdate, DESC)
        )
    )",

        @"
VAR _currdate = MAX({1})
RETURN
    CALCULATE(
        SELECTEDMEASURE(),
        FILTER(
            ALLSELECTED({1}),
            ISONORAFTER({1}, _currdate, DESC)
        )
    )",
        ordinal++,
        "Running Total, from the first date available until selected date, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

}

  if (item == "Start and End")
    {
        if (!string.IsNullOrEmpty(_FromDate) && !string.IsNullOrEmpty(_EndDate))
        {
            AddCalculationItemIfNotExists(
                "Total open - During period",
                string.Format(@"//https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = MAX({0})
VAR _StartDateVisual = MIN({0})
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        REMOVEFILTERS('{1}'),
        {2} <= _EndDateVisual,
        {3} > _StartDateVisual
        ||
        ISBLANK({3})
    )
RETURN _Result", _CalendarDate, _calendarTableName, _FromDate, _EndDate),
                ordinal++,
                string.Format("Total that have {0} and not yet {1} (total open) during the selected period", _FromDate, _EndDate),
                null
            );

            AddCalculationItemIfNotExists(
                "Open - End of period",
                string.Format(@"//https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = MAX({0})
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        REMOVEFILTERS('{1}'),
        {2} <= _EndDateVisual,
        {3} > _EndDateVisual
        ||
        ISBLANK({3})
    )
RETURN _Result", _CalendarDate, _calendarTableName, _FromDate, _EndDate),
                ordinal++,
                string.Format("Have {0} and not yet {1} (open) at the last selected Date-table date", _FromDate, _EndDate),
                null
            );
        }
        else
        {
            Warning("Start and End calculations skipped - missing required Start date or End date columns.");
        }
        
        continue;
    }


}

if (skippedItems.Count > 0)
{
    string message = "The following calculation item(s) were skipped because they already exist:\n\n" + string.Join("\n", skippedItems);
    MessageBox.Show(message, "Items Skipped", MessageBoxButtons.OK, MessageBoxIcon.Information);
}
