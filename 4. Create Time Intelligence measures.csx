//
 // Title: Create TI measures for selected measures with calc group Time Calculations
 // 
 // Author: Eivind Haugen
 // 
 // To execute this script, it will be executed based on a selected measure 
 //     and next you are promted to selected the calculation item(s):

 // It runs a nested for-loop to create all the time intelligence measures
 // 
 // For each measure with Time Intelligence, it will do the following:
 // 1. Creates the DAX expression from the selected Time Intelligence calculation item, and replaces selectedmeasure(), with the measure object (DaxObjectName)
 // 2. Adds a display-folder. 
 // 3. Apply logic for formatting. If calculation contains idx, % or pct, always set to %, if not use the format from the selected measure
 // 4. Adds a description as a combination of the description for the measure and time intelligence
 // 5. Keeps the format of the Calculation item  expression, but not formatting via DAXformatter (if it's not formatted, it's not DAX)
 // 
 // To ensure the best result, ensure that base-measures have format string and descriptions, and descriptions are included for time intelligence calculations
 //     - This is ensured if macros are run as part of this creation
 



#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

// Hide script dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// STEP 1: Get list of calculation group tables and let user select one
var calcGroupTables = Model.Tables
    .Where(t => t is CalculationGroupTable)
    .Cast<CalculationGroupTable>()
    .ToList();

if (calcGroupTables.Count == 0)
{
    MessageBox.Show("No calculation groups found in the model.");
    return;
}

var selectedCalcGroupTable = SelectTable(
    calcGroupTables,
    null,
    "Select a Calculation Group Table"
) as CalculationGroupTable;

if (selectedCalcGroupTable == null)
{
    MessageBox.Show("No calculation group table selected.");
    return;
}

// STEP 2: Let user select multiple calculation items from that group
var calcItems = selectedCalcGroupTable.CalculationItems.ToList();

Form calcItemForm = new Form();
ListBox calcItemListBox = new ListBox();
Button calcItemOkButton = new Button();
Label calcItemLabel = new Label();

calcItemForm.Text = "Select Calculation Items";
calcItemForm.Width = 350;
calcItemForm.Height = 500;

calcItemListBox.Items.AddRange(calcItems.Select(ci => ci.Name).ToArray());
calcItemListBox.SelectionMode = SelectionMode.MultiExtended;
calcItemListBox.Location = new Point(50, 20);
calcItemListBox.Width = 200;
calcItemListBox.Height = 400;

calcItemOkButton.Text = "OK";
calcItemOkButton.Location = new Point(100, 430);
calcItemOkButton.Width = 100;
calcItemOkButton.Click += (sender, e) => { calcItemForm.Close(); };

calcItemLabel.Text = "Select one or more calculation items:";
calcItemLabel.Location = new Point(50, 0);
calcItemLabel.Width = 200;

calcItemForm.Controls.Add(calcItemListBox);
calcItemForm.Controls.Add(calcItemOkButton);
calcItemForm.Controls.Add(calcItemLabel);

calcItemForm.ShowDialog();

var selectedCalcItemNames = calcItemListBox.SelectedItems.Cast<string>().ToList();
var selectedCalcItems = calcItems.Where(ci => selectedCalcItemNames.Contains(ci.Name)).ToList();

if (selectedCalcItems.Count == 0)
{
    MessageBox.Show("No calculation items selected.");
    return;
}

// STEP 3: Get already selected measures
var selectedMeasures = Selected.Measures.ToList();

if (selectedMeasures.Count == 0)
{
    MessageBox.Show("No base measures selected in Tabular Editor.");
    return;
}

// STEP 4: Validate metadata before creating new measures
// Initialize lists to track missing metadata
var missingFormatString = new List<string>();
var missingMeasureDescription = new List<string>();
var missingCalcItemDescription = new List<string>();

// Check selected measures
foreach (var m in selectedMeasures)
{
    if (string.IsNullOrWhiteSpace(m.FormatString))
        missingFormatString.Add(m.Name);
    if (string.IsNullOrWhiteSpace(m.Description))
        missingMeasureDescription.Add(m.Name);
}

// Check selected calculation items
foreach (var c in selectedCalcItems)
{
    if (string.IsNullOrWhiteSpace(c.Description))
        missingCalcItemDescription.Add(c.Name);
}

// Build info message if any metadata is missing
if (missingFormatString.Any() || missingMeasureDescription.Any() || missingCalcItemDescription.Any())
{
    string infoMessage = "Missing metadata detected and not added to new measure:\n\n" +
    "- Option 1: Press Cancel, fix items, then run script again\n" +
    "- Option 2: Apply format string/ description after creation\n";

    if (missingFormatString.Any())
    {
        infoMessage += "\n• Format string missing in measures:\n";
        foreach (var name in missingFormatString)
            infoMessage += "   - " + name + "\n";
    }

    if (missingMeasureDescription.Any())
    {
        infoMessage += "\n• Description missing in measures:\n";
        foreach (var name in missingMeasureDescription)
            infoMessage += "   - " + name + "\n";
    }

    if (missingCalcItemDescription.Any())
    {
        infoMessage += "\n• Description missing in calculation items:\n";
        foreach (var name in missingCalcItemDescription)
            infoMessage += "   - " + name + "\n";
    }

    var result = MessageBox.Show(infoMessage, "Missing Metadata", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
    if (result == DialogResult.Cancel)
        return;
}


// STEP 5: Generate new measures
string calcGroupColumnName = selectedCalcGroupTable.Columns[0].DaxObjectName;

foreach (var m in selectedMeasures)
{
    foreach (var c in selectedCalcItems)
    {
        // Choose format string. Sets % format for all measures with certain text included. If not the format of the measure
        string formatString = 
            c.Name.ToLower().Contains("idx") || 
            c.Name.Contains("%") || 
            c.Name.ToLower().Contains("pct")
            ? "#,0.0%;-#,0.0%;#,0.0%"
            : m.FormatString;

 // Use the DAX expression from the calculation item, replacing SELECTEDMEASURE() with the actual measure reference
string daxExpression;
    daxExpression = c.Expression.Replace("SELECTEDMEASURE()", m.DaxObjectName);

        // Display folder logic
        string displayFolderBase = m.DisplayFolder;
        string displayFolder = displayFolderBase + "\\" + m.Name + " - Time Intelligence";

        // Create the new measure
        var newMeasure = m.Table.AddMeasure(
            m.Name + " " + c.Name,
            daxExpression,
            displayFolder
        );

// Apply format string and description
newMeasure.FormatString = formatString;

// Only add description if both components are available
if (!string.IsNullOrWhiteSpace(m.Description) && !string.IsNullOrWhiteSpace(c.Description))
{
    newMeasure.Description = m.Description + " - using Time Intelligence definition (" + c.Name + "): " + c.Description;
}
    
        // Format DAX for readability (skipping in TE2 due to potential slow performance
        // newMeasure.Expression = FormatDax(newMeasure.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
    }
}

//MessageBox.Show("New measures created successfully.");
