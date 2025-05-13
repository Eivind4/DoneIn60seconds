
//
 // Title: Create measure Deviation (e.g. Profit/ to target)
 // 
 // Author: Eivind Haugen
 // 
 // This script, when executed, will let you create a measure with deviation and % difference between measures (only selection two measures is an option
 // 
 // The purpose of the script is to ensure best practice is included when creating measures. 
 //
 // Typical use case is to calculat Profit and Margin %, or difference to a target
 //


using System;
using System.Linq;
using System.Windows.Forms;

public class InputDialog
{
    public static string ShowInputDialog(string text, string caption, string defaultValue)
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

        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : "";
    }
}

// Hide spinner and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Input dialogs
string deviationName = InputDialog.ShowInputDialog("Provide the name for deviation measure", "Measure Name", "Profit");
string deviationPctName = InputDialog.ShowInputDialog("Provide the name for deviation % measure", "Measure Name % change", "Margin %");

// Check if exactly two measures are selected
if (Selected.Measures.Count != 2)
{
    Error("You must select exactly two measures!");
    return;
}

// Retrieve the selected measures
var selectedMeasures = Selected.Measures.ToList();
var measure1 = selectedMeasures[0];
var measure2 = selectedMeasures[1];

// Clean variable names
string var1 = "_" + measure1.Name.Replace(" ", "_");
string var2 = "_" + measure2.Name.Replace(" ", "_");
string varDev = "_deviation";

// Create DAX expressions
string daxDeviation = measure1.DaxObjectName + " - " + measure2.DaxObjectName;

string daxDeviationPct =
    "VAR " + var1 + " = " + measure1.DaxObjectName + "\n" +
    "VAR " + var2 + " = " + measure2.DaxObjectName + "\n" +
    "VAR " + varDev + " = " + var1 + " - " + var2 + "\n" +
    "RETURN DIVIDE(" + varDev + ", " + var2 + ")";

var targetTable = measure1.Table;

// Create deviation measure
if (!string.IsNullOrWhiteSpace(deviationName))
{
    var m = targetTable.AddMeasure(deviationName, daxDeviation, measure1.DisplayFolder);
    m.FormatString = measure1.FormatString;
    m.Description = "This measure is the deviation between " + measure1.DaxObjectFullName + " and " + measure2.DaxObjectFullName;
    m.Expression = FormatDax(m.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
}

// Create deviation % measure
if (!string.IsNullOrWhiteSpace(deviationPctName))
{
    var mPct = targetTable.AddMeasure(deviationPctName, daxDeviationPct, measure1.DisplayFolder);
    mPct.FormatString = "#,0.0%;-#,0.0%;#,0.0%";
    mPct.Description = "This measure is the % change between " + measure1.DaxObjectFullName + " and " + measure2.DaxObjectFullName;
    mPct.Expression = FormatDax(mPct.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
}
