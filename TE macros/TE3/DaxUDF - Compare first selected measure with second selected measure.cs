//
 // Title: Create TI measures for selected measures with Functions
 // 
 // Author: Eivind Haugen
 // 
 // To execute this script, it will be executed based on a selected measure 
 //     and next you are prompted to select the function(s):

 // It runs a nested for-loop to create all the time intelligence measures
 // 
 // For each measure with Time Intelligence, it will do the following:
 // 1. Creates the DAX expression using Function.Name + "(" + m.DaxObjectName + ")"
 // 2. Adds a display-folder. 
 // 3. Apply logic for formatting. If function name contains %, use % format, if not use the format from the selected measure
 // 4. Adds a description as a combination of the description for the measure and function
 // 5. Keeps the format of the Function expression, but not formatting via DAXformatter (if it's not formatted, it's not DAX)
 // 
 // To ensure the best result, ensure that base-measures have format string and descriptions, and descriptions are included for functions
 //     - This is ensured if macros are run as part of this creation
 

#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Improved helper function to count parameters in a function
int CountFunctionParameters(Function function)
{
    string expression = function.Expression;
    
    if (string.IsNullOrWhiteSpace(expression))
        return 0;
    
    // Find the opening parenthesis and the closing parenthesis followed by =>
    int openParenIndex = expression.IndexOf('(');
    int closeParenIndex = -1;
    
    if (openParenIndex == -1)
        return 0;
    
    // Look for ) followed by => (allowing for whitespace)
    for (int i = openParenIndex + 1; i < expression.Length - 2; i++)
    {
        if (expression[i] == ')')
        {
            // Check if this ) is followed by => (with possible whitespace)
            string remaining = expression.Substring(i + 1).TrimStart();
            if (remaining.StartsWith("=>"))
            {
                closeParenIndex = i;
                break;
            }
        }
    }
    
    if (closeParenIndex == -1)
        return 0;
    
    // Extract the parameter section between ( and )
    string parameterSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1).Trim();
    
    // If empty parameter section, return 0
    if (string.IsNullOrWhiteSpace(parameterSection))
        return 0;
    
    // Method 1: Count colons (for typed parameters like "baseMeasure: ANYREF")
    int colonCount = parameterSection.Count(c => c == ':');
    
    if (colonCount > 0)
    {
        return colonCount; // Each colon represents one parameter
    }
    
    // Method 2: Count commas and add 1 (for untyped parameters like "actual, target")
    // Remove comments first to avoid counting commas in comments
    string cleanParameterSection = RemoveComments(parameterSection);
    
    if (string.IsNullOrWhiteSpace(cleanParameterSection))
        return 0;
    
    // Count commas in the clean parameter section
    int commaCount = cleanParameterSection.Count(c => c == ',');
    
    // If there are commas, parameter count = comma count + 1
    // If no commas but content exists, it's 1 parameter
    return commaCount > 0 ? commaCount + 1 : 1;
}

// Helper function to remove comments from parameter section
string RemoveComments(string parameterSection)
{
    var lines = parameterSection.Split('\n');
    var cleanLines = new List<string>();
    
    foreach (var line in lines)
    {
        string cleanLine = line.Trim();
        
        // Skip comment-only lines (starting with //)
        if (cleanLine.StartsWith("//"))
            continue;
        
        // Remove inline comments
        int commentIndex = cleanLine.IndexOf("//");
        if (commentIndex >= 0)
        {
            cleanLine = cleanLine.Substring(0, commentIndex).Trim();
        }
        
        if (!string.IsNullOrWhiteSpace(cleanLine))
        {
            cleanLines.Add(cleanLine);
        }
    }
    
    return string.Join(" ", cleanLines);
}

// Check if exactly two measures are selected
if (Selected.Measures.Count != 2)
    throw new Exception("You must select exactly two measures!");

// Retrieve the selected measures
var measure1 = Selected.Measures.ElementAt(0);
var measure2 = Selected.Measures.ElementAt(1);

List<Function> selectedFunctions = new List<Function>();

// Check if only measures are selected (no functions)
bool onlyMeasuresSelected = Selected.Measures.Any() && !Selected.Functions.Any();

if (onlyMeasuresSelected)
{
    // STEP 1: Get list of comparison functions with exactly two parameters
    var allComparisonFunctions = Model.Functions.Where(f => f.Name.StartsWith("Comparison")).ToList();
    var twoParameterComparisonFunctions = allComparisonFunctions.Where(f => CountFunctionParameters(f) == 2).ToList();

    if (allComparisonFunctions.Count == 0)
    {
        Error("No Comparison functions found in the model.");
        return;
    }

    if (twoParameterComparisonFunctions.Count == 0)
    {
        Error($"Found {allComparisonFunctions.Count} Comparison functions, but none have exactly two input parameters.");
        return;
    }

    // STEP 2: Let user select multiple functions with improved UI
    Form functionForm = new Form();
    ListBox functionListBox = new ListBox();
    Button functionOkButton = new Button();
    Button functionCancelButton = new Button();
    Label functionLabel = new Label();

    functionForm.Text = "Select Two-Parameter Comparison Functions";
    functionForm.Width = 450;  // Increased width
    functionForm.Height = 580; // Increased height
    functionForm.StartPosition = FormStartPosition.CenterScreen;

    // Improved label with more space
    functionLabel.Text = $"Select one or more two-parameter comparison functions ({twoParameterComparisonFunctions.Count} available):";
    functionLabel.Location = new Point(20, 15);
    functionLabel.Width = 400;  // Increased width
    functionLabel.Height = 30;  // Increased height for multi-line if needed

    functionListBox.Items.AddRange(twoParameterComparisonFunctions.Select(f => f.Name).ToArray());
    functionListBox.SelectionMode = SelectionMode.MultiExtended;
    functionListBox.Location = new Point(20, 50);
    functionListBox.Width = 400;  // Increased width
    functionListBox.Height = 430; // Increased height

    // Centered and properly spaced buttons
    int buttonY = 500;  // Lower position due to increased form height
    int buttonWidth = 80;
    int buttonSpacing = 20;
    int totalButtonWidth = (buttonWidth * 2) + buttonSpacing;
    int startX = (functionForm.Width - totalButtonWidth) / 2;

    functionOkButton.Text = "OK";
    functionOkButton.Location = new Point(startX, buttonY);
    functionOkButton.Width = buttonWidth;
    functionOkButton.Height = 30;  // Increased button height
    functionOkButton.DialogResult = DialogResult.OK;

    functionCancelButton.Text = "Cancel";
    functionCancelButton.Location = new Point(startX + buttonWidth + buttonSpacing, buttonY);
    functionCancelButton.Width = buttonWidth;
    functionCancelButton.Height = 30;  // Increased button height
    functionCancelButton.DialogResult = DialogResult.Cancel;

    functionForm.Controls.Add(functionLabel);
    functionForm.Controls.Add(functionListBox);
    functionForm.Controls.Add(functionOkButton);
    functionForm.Controls.Add(functionCancelButton);

    functionForm.AcceptButton = functionOkButton;
    functionForm.CancelButton = functionCancelButton;

    var dialogResult = functionForm.ShowDialog();

    if (dialogResult == DialogResult.Cancel)
    {
        return;
    }

    var selectedFunctionNames = functionListBox.SelectedItems.Cast<string>().ToList();
    selectedFunctions = twoParameterComparisonFunctions.Where(f => selectedFunctionNames.Contains(f.Name)).ToList();
}
else
{
    if (!Selected.Functions.Any())
    {
        Error("No Functions selected.");
        return;
    }
    // Filter selected functions to only include Comparison functions with exactly two parameters
    selectedFunctions = Selected.Functions.Where(f => f.Name.StartsWith("Comparison") && CountFunctionParameters(f) == 2).ToList();
    
    if (!selectedFunctions.Any())
    {
        Error("None of the selected functions are Comparison functions with exactly two input parameters.");
        return;
    }
}

// Defensive check to prevent null reference error
if (selectedFunctions == null || selectedFunctions.Count == 0)
{
    Error("No two-parameter comparison functions selected.");
    return;
}

// Initialize lists to track missing metadata and already existing measures
var missingFormatString = new List<string>();
var missingMeasureDescription = new List<string>();
var missingFunctionDescription = new List<string>();
var alreadyExistingMeasures = new List<string>();

// Check selected measures
foreach (var m in Selected.Measures)
{
    if (string.IsNullOrWhiteSpace(m.FormatString))
        missingFormatString.Add(m.Name);
    if (string.IsNullOrWhiteSpace(m.Description))
        missingMeasureDescription.Add(m.Name);
}

// Check selected functions
foreach (var f in selectedFunctions)
{
    if (string.IsNullOrWhiteSpace(f.Description))
        missingFunctionDescription.Add(f.Name);
}

// Build info message if any metadata is missing
if (missingFormatString.Any() || missingMeasureDescription.Any() || missingFunctionDescription.Any())
{
    string infoMessage = "Missing metadata detected:\n" +
                     "- Option 1: Press Cancel, fix items, then run script again\n" +
                     "- Option 2: Continue and validate created format string and descriptions\n";

    if (missingFormatString.Any())
    {
        infoMessage += "\n• FormatString missing in measures:\n";
        foreach (var name in missingFormatString)
            infoMessage += "   - " + name + "\n";
    }

    if (missingMeasureDescription.Any())
    {
        infoMessage += "\n• Description missing in measures:\n";
        foreach (var name in missingMeasureDescription)
            infoMessage += "   - " + name + "\n";
    }

    if (missingFunctionDescription.Any())
    {
        infoMessage += "\n• Description missing in functions:\n";
        foreach (var name in missingFunctionDescription)
            infoMessage += "   - " + name + "\n";
    }

    var result = MessageBox.Show(infoMessage, "Missing Metadata", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
    if (result == DialogResult.Cancel)
    {
        return;
    }
}

// Create measures for each selected function
var finalTargetTable = measure1.Table;

foreach (var function in selectedFunctions)
{
    // Generate measure name based on function name patterns
    string measureName;
    string functionNameLower = function.Name.ToLower();
    
    if (function.Name.EndsWith("Deviation"))
    {
        measureName = measure1.Name + " Dev";
    }
    else if (functionNameLower.Contains("pct") || 
             functionNameLower.Contains("idx") || 
             functionNameLower.Contains("index"))
    {
        measureName = measure1.Name + " % change";
    }
    else
    {
        // Default fallback: remove "Comparison" prefix if present
        measureName = function.Name.Replace("Comparison", "").Trim();
        if (string.IsNullOrEmpty(measureName))
            measureName = function.Name;
    }

    // Check if measure already exists
    if (finalTargetTable.Measures.Any(m => m.Name == measureName))
    {
        alreadyExistingMeasures.Add(measureName);
        continue;
    }

    // Construct DAX expression: f.Name(measure1, measure2)
    var daxExpression = $"{function.Name}({measure1.DaxObjectName}, {measure2.DaxObjectName})";

    // Create the new measure
    var newMeasure = finalTargetTable.AddMeasure(
        measureName,
        daxExpression,
        measure1.DisplayFolder
    );

    // Determine format string
    string formatString;
    if (functionNameLower.Contains("pct") || 
        functionNameLower.Contains("idx") || 
        functionNameLower.Contains("index"))
    {
        formatString = "#,0.0%\u003B-#,0.0%\u003B#,0.0%";
    }
    else
    {
        formatString = measure1.FormatString;
    }
    newMeasure.FormatString = formatString;

    // Create description by replacing placeholders in function description
    string description = function.Description ?? $"Comparison function {function.Name}";
    description = description.Replace("measure1", measure1.DaxObjectFullName);
    description = description.Replace("measure2", measure2.DaxObjectFullName);
    newMeasure.Description = description;

    // Format the DAX expression
    newMeasure.Expression = FormatDax(newMeasure.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);
}

// Show info about already existing measures
if (alreadyExistingMeasures.Any())
{
    string existingMessage = "The following measures already exist and were skipped:\n";
    foreach (var name in alreadyExistingMeasures)
        existingMessage += "   - " + name + "\n";
    
    MessageBox.Show(existingMessage, "Already Existing Measures", MessageBoxButtons.OK, MessageBoxIcon.Information);
}