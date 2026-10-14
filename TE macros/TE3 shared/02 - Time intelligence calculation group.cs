// This script includes the select time calculations that are to be included. The idea is from the following script, but here it is included as a Calculation Group instead and with some modifications
//        Data-Goblin inspiration: https://github.com/data-goblin/powerbi-macguyver-toolbox/blob/main/tabular-editor-scripts/csharp-scripts/add-dax-templates/add-time-intelligence.csx


// There are some input that are dynamically included in the generation of time intelligence calculations:
//     1. Date column to be used from the date table (needs to be created first)
//     2. Measure that shows the last data from calendar or transaction (needs to be calculated first and have the correct formatString containing "yy"). The purpose is to hide future dates in calculations, where applicable
//     3. Based on the selected Time intelligence calculations, it is logic on which input is required. I.e. Week to date requires additional input for ISO year and ISO week column, but only if WTD is selected

// Naming is aiming at following this pattern: https://www.daxpatterns.com/standard-time-related-calculations/
// Referance to source for the expressions are included in the measure expressions

// If you know of some measures with % format, this can be updated in the logic for commonFormatString before running the script


#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;

// Don't show the script execution dialog or cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Declare table and column variables
string _timeIntelligenceCalcGroupTable = string.Empty;
string _timeIntelligenceCalcGroupColumn = string.Empty;

List<string> skippedItems = new List<string>();
List<string> createdItems = new List<string>();

// Get all Calculation Groups from the model
var calcGroups = Model.CalculationGroups.ToList();

// Check if script is run from a selected calculation group
CalculationGroupTable calculationGroupTable = null;

// Check if running from context of selected calculation group
if (Selected.CalculationGroups.Count() > 0)
{
    calculationGroupTable = Selected.CalculationGroups.First();
    _timeIntelligenceCalcGroupTable = calculationGroupTable.Name;
    _timeIntelligenceCalcGroupColumn = calculationGroupTable.Columns.Where(c => c.Name != "Ordinal").First().Name;
}
else if (calcGroups.Count > 0)
{
    // Ask user if they want to use an existing calculation group
    DialogResult dialogResult = MessageBox.Show("Do you want to add to an existing calculation group?", "Select calculation group", MessageBoxButtons.YesNo);

    if (dialogResult == DialogResult.Yes)
    {
        // Get all non-ordinal columns
        var nonOrdinalColumns = calcGroups
            .SelectMany(cg => cg.Columns.Where(c => c.Name != "Ordinal"))
            .ToList();

        if (nonOrdinalColumns.Count == 0)
        {
            MessageBox.Show("No valid Calculation Group columns found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Select the column
        string _selectedColumn = SelectColumn(nonOrdinalColumns, null, "Select Time Intelligence calc group column").DaxObjectFullName;

        // Extract table and column names
        string[] parts = _selectedColumn.Split('[');
        _timeIntelligenceCalcGroupTable = parts[0].Trim('\'');
        _timeIntelligenceCalcGroupColumn = parts[1].TrimEnd(']');

        // Get the selected Calculation Group Table
        calculationGroupTable = Model.Tables[_timeIntelligenceCalcGroupTable] as CalculationGroupTable;
    }
}

if (calculationGroupTable == null) // If user selected "No" or no calc groups exist
{
    // Prompt user for table and column names
    _timeIntelligenceCalcGroupTable = Interaction.InputBox("TABLE NAME CALCULATION GROUP:", "Calculation Group name", "Time Intelligence");
    _timeIntelligenceCalcGroupColumn = Interaction.InputBox("COLUMN NAME CALCULATION GROUP:", "Column name", "Time Calculation");

    // Create a new Calculation Group Table
    calculationGroupTable = Model.AddCalculationGroupTable(_timeIntelligenceCalcGroupTable);

    // Ensure the table was created before modifying properties
    if (calculationGroupTable != null)
    {
        (calculationGroupTable.Columns["Name"] as DataColumn).Name = _timeIntelligenceCalcGroupColumn;
        (calculationGroupTable.Columns[_timeIntelligenceCalcGroupColumn] as DataColumn).Description = "The selected time intelligence";
        calculationGroupTable.Description = "Calculation group for dynamic time intelligence.";
        calculationGroupTable.CalculationGroup.Precedence = 0;
    }
    else
    {
        MessageBox.Show("Failed to create Calculation Group Table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
}

// Define common format string
string commonFormatString = @"
    VAR MeasureInput = SELECTEDMEASURENAME()
    RETURN SWITCH(
        TRUE(),
        CONTAINSSTRING(UPPER(MeasureInput), ""DATE""), """",
        CONTAINSSTRING(MeasureInput, ""%""), ""0%"",
      //  MeasureInput = ""Margin %"", ""0%"",
        ""#,0""
    )";

// Helper method to add calculation items
void AddCalculationItem(string name, string expression, int ordinal, string description, string formatString = null)
{
    if (calculationGroupTable == null)
    {
        MessageBox.Show("Calculation Group Table is not initialized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }

    var item = calculationGroupTable.AddCalculationItem(name);
    item.Expression = expression;
    item.FormatStringExpression = formatString ?? commonFormatString;
    item.Ordinal = ordinal;
    item.Description = description;
}

List<string> _ToDateOptions = new List<string> { "Actual","MTD", "QTD", "YTD", "R12", "Rolling average", "Rolling total", "YTD FY", "WTD", "IB", "Start/End", "Last Date" };

// Create a new form
var _Form = new Form()
{
Text = "Add which time intelligence categories?",
Padding = new Padding(5),  // 5 px padding around the edges
StartPosition = FormStartPosition.CenterParent,
FormBorderStyle = FormBorderStyle.FixedDialog,
AutoSize = true,
AutoSizeMode = AutoSizeMode.GrowAndShrink,
MaximizeBox = false,
MinimizeBox = false,
ShowIcon = false,
MinimumSize = new Size(400, 0)
};

// Create checkboxes
var _Checkbox1 = new CheckBox()
{
Text = "Actual",
AutoSize = true,
Checked = true
};

var _Checkbox2 = new CheckBox()
{
Text = "MTD",
AutoSize = true,
Checked = true
};

var _Checkbox3 = new CheckBox()
{
Text = "QTD",
AutoSize = true
};

var _Checkbox4 = new CheckBox()
{
Text = "YTD",
AutoSize = true,
Checked = true
};

var _Checkbox5 = new CheckBox()
{
Text = "R12",
AutoSize = true,
Checked = true
};

var _Checkbox6 = new CheckBox()
{
Text = "Rolling average",
AutoSize = true,
Checked = true
};

var _Checkbox7 = new CheckBox()
{
Text = "Running total",
AutoSize = true
};

var _Checkbox8 = new CheckBox()
{
    Text = "YTD FY",
    AutoSize = true
};

var _Checkbox9 = new CheckBox()
{
    Text = "WTD",
    AutoSize = true
};

var _Checkbox10 = new CheckBox()
{
    Text = "IB",
    AutoSize = true
};

var _Checkbox11 = new CheckBox()
{
    Text = "Start/End",
    AutoSize = true
};

var _Checkbox12 = new CheckBox()
{
    Text = "Last Date",
    AutoSize = true
};

// Arrange checkboxes vertically
var _CheckBoxPanel = new FlowLayoutPanel()
{
FlowDirection = FlowDirection.TopDown,
AutoSize = true,
AutoSizeMode = AutoSizeMode.GrowOnly,
Dock = DockStyle.Top,
Padding = new Padding(0, 0, 0, 10), // Padding bottom
WrapContents = false
};

// Add checkboxes to the panel
_CheckBoxPanel.Controls.Add(_Checkbox1);
_CheckBoxPanel.Controls.Add(_Checkbox2);
_CheckBoxPanel.Controls.Add(_Checkbox3);
_CheckBoxPanel.Controls.Add(_Checkbox4);
_CheckBoxPanel.Controls.Add(_Checkbox5);
_CheckBoxPanel.Controls.Add(_Checkbox6);
_CheckBoxPanel.Controls.Add(_Checkbox7);
_CheckBoxPanel.Controls.Add(_Checkbox8);
_CheckBoxPanel.Controls.Add(_Checkbox9);
_CheckBoxPanel.Controls.Add(_Checkbox10);
_CheckBoxPanel.Controls.Add(_Checkbox11);
_CheckBoxPanel.Controls.Add(_Checkbox12);

// Create panel for buttons
var _ButtonPanel = new FlowLayoutPanel()
{
FlowDirection = FlowDirection.RightToLeft,
AutoSize = true,
AutoSizeMode = AutoSizeMode.GrowAndShrink,
Dock = DockStyle.Bottom,
Padding = new Padding(0, 10, 0, 0) // Padding top
};

// Create OK and Cancel buttons
var _OkButton = new Button()
{
Text = "Create Time intelligence",
DialogResult = DialogResult.OK,
AutoSize = true
};

var _CancelButton = new Button()
{
Text = "Cancel",
DialogResult = DialogResult.Cancel,
AutoSize = true
};

// Add buttons to the panel
_ButtonPanel.Controls.Add(_CancelButton);
_ButtonPanel.Controls.Add(_OkButton);

// Set form's Accept and Cancel buttons
_Form.AcceptButton = _OkButton;
_Form.CancelButton = _CancelButton;

// Add panels to the form
_Form.Controls.Add(_CheckBoxPanel);
_Form.Controls.Add(_ButtonPanel);

// Display the form and capture the result
var _Result = _Form.ShowDialog();

// Creates a list off columns of type date that is marked as date in a date-table
var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime &&
    c.IsKey == true).ToList();

// Creates a list of measures that have a date format (date-type format string)
var _dateFormattedMeasures = Model.AllMeasures
    .Where(measure => measure.FormatString != null && measure.FormatString.Contains("yy"))
    .ToList();

//--------------------Input that is required for all time intelligence calculations-------------------------
// Select the date column from calendar table to be used in calculation items
try
{
    // Select the date-key in the model to be used in time-intelligence calculations
    string _CalendarDate = SelectColumn(_dateColumns, null, "GENERAL INPUT:" + Environment.NewLine +"Select the DATE-COLUMN to be used" + Environment.NewLine + "for time-calculations").DaxObjectFullName;
    
    // Extract the calendar table for the selected date column
    var _calendarTableName = _CalendarDate.Split('[')[0].Trim('\'');

// ----------------Checkbox values start-----------------------
if (_Result == DialogResult.OK)
{
    // Retrieve checkbox values
    bool _ActualChecked = _Checkbox1.Checked;
    bool _MtdChecked = _Checkbox2.Checked;
    bool _QtdChecked = _Checkbox3.Checked;
    bool _YtdChecked = _Checkbox4.Checked;
    bool _R12Checked = _Checkbox5.Checked;
    bool _RaverageChecked = _Checkbox6.Checked;
    bool _RtotalChecked = _Checkbox7.Checked;
    bool _YtdFYChecked = _Checkbox8.Checked;
    bool _WtdChecked = _Checkbox9.Checked;
    bool _IBChecked = _Checkbox10.Checked;
    bool _StartEndChecked = _Checkbox11.Checked;
    bool _LastDateChecked = _Checkbox12.Checked;

//---------------To check if the calculation item is already present------------
// Helper function to check if a calculation item exists
bool CalculationItemExists(string itemName)
{
    return calculationGroupTable.CalculationItems.Any(ci => ci.Name == itemName);
}

// Function to add a calculation item only if it doesn't already exist
Action<string, string, int, string, string> AddCalculationItemIfNotExists = 
(name, expression, ordinalPos, description, formatString) =>
{
    if (calculationGroupTable.CalculationItems.Any(ci => ci.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase)))
    {
        skippedItems.Add(name);
        return;
    }

    var item = calculationGroupTable.AddCalculationItem(name);
    item.Expression = expression;
    item.Ordinal = ordinalPos;
    item.Description = description;
    
    // Apply the format string
    if (!string.IsNullOrEmpty(formatString))
    {
        item.FormatStringExpression = formatString;
    }
    else
    {
        item.FormatStringExpression = commonFormatString;
    }
    
    createdItems.Add(name);
};


        // Track ordinal position - start from highest existing ordinal + 1
        int ordinal = 1;
        if (calculationGroupTable.CalculationItems.Count > 0)
        {
            ordinal = calculationGroupTable.CalculationItems.Max(ci => ci.Ordinal) + 1;
        }



//-------------------------------START Time Intelligence Calculations where hiding future dates are included---------
if (_ActualChecked || _MtdChecked || _QtdChecked || _YtdChecked || _R12Checked || _RaverageChecked || _RtotalChecked || _YtdFYChecked)  
{ 
    try
    {
        // Show dialog to choose between Column and Measure approach
        Form futureLogicForm = new Form();
        futureLogicForm.Text = "Hide Future Dates";
        futureLogicForm.Width = 500;  // Increased from 450
        futureLogicForm.Height = 300; // Increased from 250
        futureLogicForm.FormBorderStyle = FormBorderStyle.FixedDialog;
        futureLogicForm.StartPosition = FormStartPosition.CenterScreen;

        Label promptLabel = new Label()
        {
            Left = 20,
            Top = 20,
            Width = 450,  // Increased from 400
            Height = 20,
            Text = "Choose method for hiding future dates:"
        };

        Label columnInfoLabel = new Label()
        {
            Left = 20,
            Top = 45,
            Width = 450,  // Increased from 400
            Height = 40,   // Increased from 30
            Text = "• Column: Requires column in date-table to filter dates (i.e. Is Historical (True/False). Likely to be more performant",
            AutoSize = false
        };

        Label measureInfoLabel = new Label()
        {
            Left = 20,
            Top = 85,     // Adjusted position
            Width = 450,  // Increased from 400
            Height = 40,  // Increased from 30
            Text = "• Measure: Select a measure (Measure preferred, but not required). More flexible if no column to filter dates",
            AutoSize = false
        };

        Button columnButton = new Button()
        {
            Text = "Column",
            Left = 70,    // Adjusted position
            Width = 120,  // Increased from 100
            Height = 35,  // Added height
            Top = 140,    // Adjusted position
            DialogResult = DialogResult.Yes
        };

        Button measureButton = new Button()
        {
            Text = "Measure",
            Left = 200,   // Adjusted position
            Width = 120,  // Increased from 100
            Height = 35,  // Added height
            Top = 140,    // Adjusted position
            DialogResult = DialogResult.No
        };

        Button cancelButton = new Button()
        {
            Text = "Cancel",
            Left = 330,   // Adjusted position
            Width = 120,  // Increased from 100
            Height = 35,  // Added height
            Top = 140,    // Adjusted position
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

        string futureDateFilterType = (futureResult == DialogResult.Yes) ? "Column" : "Measure";
        string _LastDateAvailable = string.Empty;

        if (futureDateFilterType == "Column")
        {
            // Get columns from the date table
            var dateTable = Model.Tables[_calendarTableName];
            var dateTableColumns = dateTable.Columns.ToList();
            
            var selectedColumn = SelectColumn(dateTableColumns, null, "Select the COLUMN to filter future dates" + Environment.NewLine + "(e.g., Is Historical)");
            
            if (selectedColumn == null)
            {
                Error("No column selected for filtering future dates.");
                return;
            }
            
            // Ask user for the column value to filter by
            string columnValue = Interaction.InputBox(
                "Enter the VALUE that the column should equal to show historical dates:" + Environment.NewLine + 
                Environment.NewLine +
                "Examples:" + Environment.NewLine +
                "• TRUE (for boolean columns)" + Environment.NewLine +
                "• 1 (for integer columns)" + Environment.NewLine +
                "• 'Yes' (for text columns - include quotes)",
                "Column Filter Value",
                "TRUE"
            );
            
            if (string.IsNullOrEmpty(columnValue))
            {
                Error("No column value provided for filtering future dates.");
                return;
            }
            
            _LastDateAvailable = selectedColumn.DaxObjectFullName + " = " + columnValue;
        }
        else
        {
            // Select the measure to be used for hiding future dates
            var selectedMeasure = SelectMeasure(_dateFormattedMeasures, null, "INPUT HIDE FUTURE DATES:" + Environment.NewLine +
                "Select the measure to be used for:" + Environment.NewLine + 
                "      - Hiding future dates " + Environment.NewLine + 
                "(require measure with yy or yyyy format string)"+ Environment.NewLine +
                " "+ Environment.NewLine +
                "CANCEL option:"+ Environment.NewLine + 
                 "        Sets it to min of last calendar date and today");

            _LastDateAvailable = selectedMeasure != null 
                ? $"{selectedMeasure.DaxObjectName}" 
                : $"MIN(MAX({_CalendarDate}), TODAY())";
        }

        // Helper function for time intelligence items with future date hiding
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

        // Get selected average for rolling calculations if needed
        string _selectedAverage = string.Empty;
        if (_RaverageChecked)
        {
            try
            {
                var _selectedTableName = _CalendarDate.Split('[')[0].Trim('\'');
                var _selectedDateTable = Model.Tables[_selectedTableName];
                var _columnsInDateTable = _selectedDateTable.Columns.ToList();
                _selectedAverage = SelectColumn(_columnsInDateTable, null, "Select the column to be used as average " + Environment.NewLine + "in rolling averages (ie. daily, monthly)").DaxObjectFullName;
            }
            catch
            {
                Error("No column selected for rolling averages");
                return;
            }
        }



 // ----------------------START - Actual
if (_ActualChecked)
{
    AddCalculationItemIfNotExists(
        "Actual",
        "SELECTEDMEASURE()",
        ordinal++,
        "Actual value",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "Actual to date",
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
        ordinal++,
        "Actual value, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddCalculationItemIfNotExists(
        "Actual Y-1",
        string.Format(@"CALCULATE(
            SELECTEDMEASURE(),
            SAMEPERIODLASTYEAR({0})
        )", _CalendarDate),
        ordinal++,
        "Actual value last year",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "Actual to date Y-1",
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
        ordinal++,
        "Actual value last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "AOA",
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
        ordinal++,
        "AOA=Actual over Actual: Deviation for actual value this year and last year, hiding future dates {FILTER_DESCRIPTION}",
        null
    );

    AddTimeIntelligenceItemWithHideFuture(
        "AOA %",
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
        ordinal++,
        "AOA=Actual over Actual: % change this year and last year, hiding future dates {FILTER_DESCRIPTION}",
        "\"0%\""
    );

    AddCalculationItemIfNotExists(
        "AOA C",
        string.Format(@"VAR _thisyear = SELECTEDMEASURE()
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR({0})
    )
RETURN _thisyear - _lastyear", _CalendarDate),
        ordinal++,
        "AOA C=Actual over Actual Complete: Deviation for actual value this year and last year, not hiding future dates",
        null
    );

    AddCalculationItemIfNotExists(
        "AOA C %",
        string.Format(@"VAR _thisyear = SELECTEDMEASURE()
VAR _lastyear =
    CALCULATE(
        SELECTEDMEASURE(),
        SAMEPERIODLASTYEAR({0})
    )
VAR _deviationYear = _thisyear - _lastyear
RETURN DIVIDE(_deviationYear, _lastyear)", _CalendarDate),
        ordinal++,
        "AOA C=Actual over Actual Complete: % change this year and last year, not hiding future dates",
        "\"0%\""
    );
}  //END - Actual

        // ----------------------START - MTD
        if (_MtdChecked)
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

    if (_QtdChecked)
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

    if (_YtdChecked)
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

    if (_R12Checked)  
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

    if (_RaverageChecked)
    {
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

     AddRollingAverageItem(
        "Rolling 12M avg",
        @"// Reference: https://www.sqlbi.com/articles/rolling-12-months-average-in-dax/
//Calculating the average based on the selection in VALUES, such as per day or per month
VAR _NumOfMonths = 12
VAR _LastCurrentDate = CALCULATE(
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
//Calculating the average based on the selection in VALUES, such as per day or per month
VAR _LastDayShown = {0}
VAR _NumOfMonths = 12
VAR _LastCurrentDate = CALCULATE(
    MAX({1}), FILTER(
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
        null
    );

}

    if (_RtotalChecked)
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


    // ----------------------START - YTD FY (with HideFuture)
    if (_YtdFYChecked)
    {
        try
        {
            // Sets the end day for the fiscal year
            string _fiscalYearInput = Interaction.InputBox("FISCAL YEAR INPUT:" + Environment.NewLine +"Last day of the fiscal year", "Fiscal year input, last date MM/DD", "06/30");    

            if (string.IsNullOrEmpty(_fiscalYearInput))
            {
                Error("No fiscal year input provided.");
                return;
            }

   AddTimeIntelligenceItemWithHideFuture(
    "YTD FY",
    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({{1}}),
        {{0}}
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
RETURN _Result", _fiscalYearInput),

    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {{0}} 
VAR _CurrentDates = 
    FILTER(
        VALUES({{1}}),
        {{1}} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
RETURN _Result", _fiscalYearInput),
    ordinal++,
    "Accumulated year to date for the fiscal year, hiding future dates {FILTER_DESCRIPTION}",
    null
);

// LYTD FY
AddTimeIntelligenceItemWithHideFuture(
    "LYTD FY",
    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({{1}}),
        {{0}}
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
RETURN _Result", _fiscalYearInput),

    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {{0}} 
VAR _CurrentDates = 
    FILTER(
        VALUES({{1}}),
        {{1}} <= _LastDayAvailable
    )
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
RETURN _Result", _fiscalYearInput),
    ordinal++,
    "Accumulated year to date for fiscal year last year, hiding future dates {FILTER_DESCRIPTION}",
    null
);

// YOYTD FY
AddTimeIntelligenceItemWithHideFuture(
    "YOYTD FY",
    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({{1}}),
        {{0}}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
RETURN _CurrentYear - _PreviousYear", _fiscalYearInput),

    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {{0}} 
VAR _CurrentDates = 
    FILTER(
        VALUES({{1}}),
        {{1}} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
RETURN _CurrentYear - _PreviousYear", _fiscalYearInput),
    ordinal++,
    "YOYTD FY= Year over year to date Fiscal Year: Year-to-Date fiscal year Deviation, this year and last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
    null
);

// YOYTD FY %
AddTimeIntelligenceItemWithHideFuture(
    "YOYTD FY %",
    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _CurrentDates = 
    CALCULATETABLE(
        VALUES({{1}}),
        {{0}}
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)", _fiscalYearInput),

    string.Format(@"// https://www.sqlbi.com/articles/hiding-future-dates-for-calculations-in-dax/
VAR _LastDayAvailable = {{0}} 
VAR _CurrentDates = 
    FILTER(
        VALUES({{1}}),
        {{1}} <= _LastDayAvailable
    )
VAR _CurrentYear =
    CALCULATE (
        SELECTEDMEASURE (),
        DATESYTD (_CurrentDates, ""{0}"")
    )
VAR _PreviousYear =
    CALCULATE(
        SELECTEDMEASURE(),
        DATESYTD(
            SAMEPERIODLASTYEAR(_CurrentDates), ""{0}""
        )
    )
VAR _DeltaYear = _CurrentYear - _PreviousYear
RETURN DIVIDE(_DeltaYear, _PreviousYear)", _fiscalYearInput),
    ordinal++,
    "YOYTD FY= Year over year to date Fiscal Year: Year-to-Date fiscal year Index, % change this year compared to last year accumulated, hiding future dates {FILTER_DESCRIPTION}",
    "\"0%\""
);
        }
        catch
        {
            Error("No fiscal year input. Time intelligence calculations not created");
        }
    }   // END YTD FY


    }    
    catch
    {
        Error("Error in future date logic selection.");
    }
}    
//---------------------------END Calculation items with hiding future dates--------------------------------

    // ----------------------START - WTD
    if (_WtdChecked)
    {
        // Create a variable with all columns from the calendar table
        var _calendarColumns = Model.AllColumns
            .Where(c => string.Equals(c.Table.Name, _calendarTableName, StringComparison.OrdinalIgnoreCase) 
                        && (c.DataType == DataType.Int64 || c.DataType == DataType.Double)) // Only numeric columns
            .ToList();

        try
        {
            // Select the ISO-year in the model to be used in week time-intelligence calculations
            string _ISO_year = SelectColumn(_calendarColumns, null, "WEEK TO DATE CALCULATIONS:" + Environment.NewLine +"Select the ISO_YEAR-COLUMN to be used" + Environment.NewLine + "(assumption: created in calendar-table)").DaxObjectFullName;
            // Select the ISO-week in the model to be used in week time-intelligence calculations
            string _ISO_week = SelectColumn(_calendarColumns, null, "WEEK TO DATE CALCULATIONS:" + Environment.NewLine +"Select the ISO_Week-COLUMN (WW-format) to be used" + Environment.NewLine + "(assumption: created in calendar-table)").DaxObjectFullName;
            // Select the day number in the model to be used in week time-intelligence calculations
            string _dayNoOfWeek = SelectColumn(_calendarColumns, null, "WEEK TO DATE CALCULATIONS:" + Environment.NewLine +"Select the DAY NUMBER COLUMN to be used" + Environment.NewLine + "(assumption: created in calendar-table)").DaxObjectFullName;

            AddCalculationItemIfNotExists(
                "WTD",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} )
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
        )
    ),
    BLANK ()
)", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate),
                ordinal++,
                "Accumulated week to date for ISO week",
                null
            );

            AddCalculationItemIfNotExists(
                "WTD LY",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} ) - 1
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
               && {4} <= MAX ( {4} )
        )
    ),
    BLANK ()
)", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "Accumulated week to date same ISO week last year",
                null
            );

             AddCalculationItemIfNotExists(
                "WOWTD",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
VAR _thisYear =
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} )
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
        )
    ),
    BLANK ()
)

VAR _lastYear = 
   IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} ) - 1
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
               && {4} <= MAX ( {4} )
        )
    ),
    BLANK ()
)
VAR _Result = _thisYear-_lastYear
RETURN _Result", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "WOWTD = Week over week to date (last year). Accumulated week to date deviation to corresponding ISO week last year",
                null
            );

            AddCalculationItemIfNotExists(
                "WOWTD %",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
VAR _thisYear =
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} )
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
        )
    ),
    BLANK ()
)

VAR _lastYear = 
   IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} ) - 1
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
               && {4} <= MAX ( {4} )
        )
    ),
    BLANK ()
)
VAR _Result = DIVIDE(_thisYear-_lastYear,_lastYear)
RETURN _Result", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "WOWTD % = Week over week to date (Last year). % Change for accumulated week to date to corresponding ISO week last year",
                "\"0%\""
            );

            AddCalculationItemIfNotExists(
                "PWTD",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
VAR _CurrentYear = MAX({0})
VAR _CurrentWeek = MAX({1})
VAR _thePreviousYear = _CurrentYear - 1
VAR _LastWeekPrevYear =
    CALCULATE (
    MAX ( {1} ),
    ALL ( '{2}' ),  -- Remove existing filters
        {0} = _thePreviousYear
    )

VAR _TargetYear = IF ( _CurrentWeek = 1, _thePreviousYear, _CurrentYear )
VAR _TargetWeek = IF ( _CurrentWeek = 1, _LastWeekPrevYear, _CurrentWeek - 1 )

VAR _Result = 
    IF (
HASONEVALUE ({0})
    && HASONEVALUE ({1} ),
CALCULATE (
   SELECTEDMEASURE(),
    FILTER (
        ALL ( '{2}' ),
        {0} = _TargetYear
            && {1} = _TargetWeek
            && {3} <= MAX ( {3} )
            && {4}  <= MAX ( {4}  )
    )
),
BLANK ()
)

RETURN _Result", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "Accumulated week to date previous ISO week",
                null
            );

            AddCalculationItemIfNotExists(
                "WOPWTD",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
VAR _thisYear =
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} )
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
        )
    ),
    BLANK ()
)

// Last week
VAR _CurrentYear = MAX({0})
VAR _CurrentWeek = MAX({1})
VAR _thePreviousYear = _CurrentYear - 1
VAR _LastWeekPrevYear =
    CALCULATE (
    MAX ( {1} ),
    ALL ( '{2}' ),  -- Remove existing filters
        {0} = _thePreviousYear
    )

VAR _TargetYear = IF ( _CurrentWeek = 1, _thePreviousYear, _CurrentYear )
VAR _TargetWeek = IF ( _CurrentWeek = 1, _LastWeekPrevYear, _CurrentWeek - 1 )

VAR _lastWeek =  
    IF (
HASONEVALUE ({0})
    && HASONEVALUE ({1} ),
CALCULATE (
   SELECTEDMEASURE(),
    FILTER (
        ALL ( '{2}' ),
        {0} = _TargetYear
            && {1} = _TargetWeek
            && {3} <= MAX ( {3} )
            && {4}  <= MAX ( {4}  )
    )
),
BLANK ()
)
VAR _Result = _thisYear-_lastWeek
RETURN _Result", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "WOPWTD = Week over previous week to date: Accumulated week to date difference to previous ISO week",
                null
            );

            AddCalculationItemIfNotExists(
                "WOPWTD %",
                string.Format(@"// https://www.sqlbi.com/articles/week-based-time-intelligence-in-dax
// Shows the entire week without hiding future dates
VAR _thisYear =
IF (
    HASONEVALUE ({0})
        && HASONEVALUE ({1} ),
    CALCULATE (
       SELECTEDMEASURE(),
        FILTER (
            ALL ( '{2}' ),
            {0} = VALUES ( {0} )
                && {1} = VALUES ( {1} )
               && {3} <= MAX ( {3} )
        )
    ),
    BLANK ()
)

//Last week
VAR _CurrentYear = MAX({0})
VAR _CurrentWeek = MAX({1})
VAR _thePreviousYear = _CurrentYear - 1
VAR _LastWeekPrevYear =
    CALCULATE (
    MAX ( {1} ),
    ALL ( '{2}' ),  -- Remove existing filters
        {0} = _thePreviousYear
    )

VAR _TargetYear = IF ( _CurrentWeek = 1, _thePreviousYear, _CurrentYear )
VAR _TargetWeek = IF ( _CurrentWeek = 1, _LastWeekPrevYear, _CurrentWeek - 1 )

VAR _lastWeek = 
    IF (
HASONEVALUE ({0})
    && HASONEVALUE ({1} ),
CALCULATE (
   SELECTEDMEASURE(),
    FILTER (
        ALL ( '{2}' ),
        {0} = _TargetYear
            && {1} = _TargetWeek
            && {3} <= MAX ( {3} )
            && {4}  <= MAX ( {4}  )
    )
),
BLANK ()
)
VAR _Result = DIVIDE(_thisYear-_lastWeek,_lastWeek)
RETURN _Result", _ISO_year, _ISO_week, _calendarTableName, _CalendarDate, _dayNoOfWeek),
                ordinal++,
                "WOPWTD % = Week over previous week to date %: Accumulated week to date index: % change from previous ISO week",
                "\"0%\""
            );
            
        }
        catch
        {
            Error("No ISO-week input. Time intelligence calculations not created");
        }
    }

    // ----------------------START - IB
    if (_IBChecked)
    {
        AddCalculationItemIfNotExists(
            "IB",
            string.Format(@"VAR _earliestDate = MIN({0})
RETURN
    CALCULATE(
        SELECTEDMEASURE(),
        FILTER(
            ALLSELECTED({0}),
            ISONORAFTER({0}, _earliestDate, DESC)
        ),
            {0}<_earliestDate
    )", _CalendarDate),
            ordinal++,
            "Running Total, before the first selected calendar date",
            null
        );
    }

    // ----------------------START - Start/End Date
    if (_StartEndChecked)
    {
        // Creates a list off columns of type date that is marked as date in a date-table
        var _dateColumnsNotCalendar = Model.AllColumns.Where(c => c.DataType == DataType.DateTime &&
            c.IsKey != true).ToList();

        try
        {
            // Select the Start date (not the date in the calendar)
            string _FromDate = SelectColumn(_dateColumnsNotCalendar, null, "START AND END DATE CALCULATIONS:" + Environment.NewLine +"Select the START-date column to be used" + Environment.NewLine + "(Assumption: DateTime formatted)").DaxObjectFullName;

            // Select the End date (not the date in the calendar)
            string _EndDate = SelectColumn(_dateColumnsNotCalendar, null, "START AND END DATE CALCULATIONS:" + Environment.NewLine +"Select the END-date column to be used" + Environment.NewLine + "(Assumption: DateTime formatted)").DaxObjectFullName;

            AddCalculationItemIfNotExists(
                "Total open - During period",
                string.Format(@"//https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = MAX({0})
VAR _StartDateVisual = MIN({0})
VAR _Result =
    Calculate(
        SELECTEDMEASURE(),
        REMOVEFILTERS('{1}'),
        {2}<= _EndDateVisual,
        {3}> _StartDateVisual
        ||
        ISBLANK({3})
    )
Return _Result", _CalendarDate, _calendarTableName, _FromDate, _EndDate),
                ordinal++,
                string.Format("Total that have {0} and not yet {1} (total open) during the selected period", _FromDate, _EndDate),
                null
            );

            AddCalculationItemIfNotExists(
                "Open - End of period",
                string.Format(@"//https://www.youtube.com/watch?v=YL7H1Rqckb0
VAR _EndDateVisual = max({0})
VAR _Result =
    Calculate(
        SELECTEDMEASURE(),
        REMOVEFILTERS('{1}'),
        {2}<= _EndDateVisual,
        {3}> _EndDateVisual
        ||
        ISBLANK({3})
    )
Return _Result", _CalendarDate, _calendarTableName, _FromDate, _EndDate),
                ordinal++,
                string.Format("Have {0} and not yet {1} (open) at the last selected Date-table date", _FromDate, _EndDate),
                null
            );
        }   
        catch
        {
            Error("No Start and end-date selected. Time intelligence calculation items not created");
        }
    }

    // ----------------------START - Last Date
    if (_LastDateChecked)
    {
        // Creates a list off columns of type date that is marked as date in a date-table
        var _dateColumnsNotCalendarBalance = Model.AllColumns.Where(c => c.DataType == DataType.DateTime &&
            c.IsKey != true).ToList();

        try
        {
            // Select the End date (not the date in the calendar)
            string _balanceDate = SelectColumn(_dateColumnsNotCalendarBalance, null, "LAST DATE (I.E. BALANCE, LAST N DAYS):" + Environment.NewLine +"Select the fact table date column to be used" + Environment.NewLine + "(Assumption: DateTime formatted)").DaxObjectFullName;

            // Extract the fact table for the selected date column
            var _balanceTableName = _balanceDate.Split('[')[0].Trim('\'');

        AddCalculationItemIfNotExists(
            "LastDate",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxBalanceDate =
    CALCULATE (
        MAX ( {0} ),  // Use MIN for Balance FirstDateWithData
        ALLEXCEPT (
            '{1}',            // Remove filters from the Balances expanded table
            '{2}'               // but not from the date
        )
    )
VAR _Result =
    CALCULATE (
        Selectedmeasure(),
        {3} = _MaxBalanceDate
    )
RETURN _Result", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "Last date in a fact table, i.e. a balance",
        null
        );

        AddCalculationItemIfNotExists(
            "LastDate per Category",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})  // Get the max date from the Date table
VAR _noDays = 1   // [Rn input] if using N last date from selected value input parameter
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )
RETURN _Result", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "Last date for each category with value",
        null
        );

        AddCalculationItemIfNotExists(
            "LastDate per Category LY",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _MaxFactDateLY = DATE(YEAR(_MaxFactDate)-1,MONTH(_MaxFactDate),DAY(_MaxFactDate))
VAR _noDays = 1
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDateLY,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )
RETURN _Result", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "Last date for each category with value Last year",
        null
        );

        AddCalculationItemIfNotExists(
            "LastDate per Category LY Dev",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _MaxFactDateLY = DATE(YEAR(_MaxFactDate)-1,MONTH(_MaxFactDate),DAY(_MaxFactDate))
VAR _noDays = 1

// Last period
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _ResultLast =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )

// Last period last year
VAR _ProdDatesLY =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDateLY,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDatesLY =
    TOPN(_noDays, _ProdDatesLY, {0}, DESC)
VAR _ResultLY =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDatesLY
    )
RETURN _ResultLast - _ResultLY", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "Actual change, Last date for each category with value Last year",
        null
        );

        AddCalculationItemIfNotExists(
            "LastDate per Category LY idx",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _MaxFactDateLY = DATE(YEAR(_MaxFactDate)-1,MONTH(_MaxFactDate),DAY(_MaxFactDate))
VAR _noDays = 1

// Last period
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _ResultLast =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )

// Last period last year
VAR _ProdDatesLY =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDateLY,
        REMOVEFILTERS('{2}')
    )
VAR _LastProdDatesLY =
    TOPN(_noDays, _ProdDatesLY, {0}, DESC)
VAR _ResultLY =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDatesLY
    )

VAR _ResultDeviation = _ResultLast - _ResultLY
RETURN DIVIDE(_ResultDeviation,_ResultLY)", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "% change Last date for each category with value Last year",
            "\"0%\""
        );

        AddCalculationItemIfNotExists(
            "Previous per Category",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _noDays = 1
VAR _noDaysPrevious = _noDays*2
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _PreviousProdDates =
    TOPN(_noDaysPrevious, _ProdDates, {0}, DESC)
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _Result =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _PreviousProdDates,
        NOT {3} IN _LastProdDates
    )
RETURN _Result", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "The previous period for category. I.e. if compared to last date, then it is previous date with value. If 7 days, then the 7 days before with values",
        null
        );

        AddCalculationItemIfNotExists(
            "Previous per Category LY Dev",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _noDays = 1
VAR _noDaysPrevious = _noDays*2
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _PreviousProdDates =
    TOPN(_noDaysPrevious, _ProdDates, {0}, DESC)
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _ResultPrevious =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _PreviousProdDates,
        NOT {3} IN _LastProdDates
    )
VAR _ResultsLast =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )
RETURN _ResultsLast - _ResultPrevious", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "Actual change, The previous period for category. I.e. if compared to last date, then it is previous date with value. If 7 days, then the 7 days before with values",
        null
        );

        AddCalculationItemIfNotExists(
            "Previous per Category LY idx",
            string.Format(@"
// https://www.daxpatterns.com/semi-additive-calculations/
VAR _MaxFactDate = MAX({0})
VAR _noDays = 1
VAR _noDaysPrevious = _noDays*2
VAR _ProdDates =
    CALCULATETABLE(
        VALUES({0}),
        {0} <= _MaxFactDate,
        REMOVEFILTERS('{2}')
    )
VAR _PreviousProdDates =
    TOPN(_noDaysPrevious, _ProdDates, {0}, DESC)
VAR _LastProdDates =
    TOPN(_noDays, _ProdDates, {0}, DESC)
VAR _ResultPrevious =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _PreviousProdDates,
        NOT {3} IN _LastProdDates
    )
VAR _ResultsLast =
    CALCULATE(
        SELECTEDMEASURE(),
        {3} IN _LastProdDates
    )
VAR _ResultsDeviation = _ResultsLast - _ResultPrevious
RETURN DIVIDE(_ResultsDeviation,_ResultPrevious)", _balanceDate, _balanceTableName, _calendarTableName, _CalendarDate),
            ordinal++,
            "% change The previous period for category. I.e. if compared to last date, then it is previous date with value. If 7 days, then the 7 days before with values",
            "\"0%\""
        );

            // Add more LastDate calculations following the same pattern...
        }
        catch
        {
            Error("No Last balance date calculation items created");
        }
    }





//---------------------------END checkbox values ---------------------------------
}


// Show summary of what was created and skipped
if (createdItems.Count > 0 || skippedItems.Count > 0)
{
    string summary = "Script execution completed:\n\n";
    
    if (createdItems.Count > 0)
    {
        summary += $"Created {createdItems.Count} calculation items:\n";
        foreach (var item in createdItems)
        {
            summary += $"  • {item}\n";
        }
        summary += "\n";
    }
    
    if (skippedItems.Count > 0)
    {
        summary += $"Skipped {skippedItems.Count} existing calculation items:\n";
        foreach (var item in skippedItems)
        {
            summary += $"  • {item}\n";
        }
    }
    
    MessageBox.Show(summary, "Time Intelligence Script Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
}


//---------------------------END Defining input parameter: Date from calendar table-------------------------------------
}
catch
{
    Error("No selection of calendar marked as date-table with date-column! " + Environment.NewLine + "Ending script without creating calculation items.");
}