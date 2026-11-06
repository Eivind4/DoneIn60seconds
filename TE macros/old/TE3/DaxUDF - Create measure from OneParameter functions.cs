// ================================================================================================
// Title: Create Time Intelligence Measures for Selected Measures with Functions
// Author: Eivind Haugen
// ================================================================================================
//
// DESCRIPTION:
// This script automates the creation of time intelligence (or other calculated) measures by 
// applying single-parameter functions to selected base measures. It generates new measures with
// properly formatted names, display folders, and metadata based on function annotations.
//
// HOW TO USE:
// 1. Select one or more base measures in your model
// 2. Run the script
// 3. Select one or more single-parameter functions from the dialog (functions starting with 
//    "Local" will be shown)
// 4. If functions are missing annotations, you'll be prompted to provide measure and folder 
//    suffixes (once per function)
// 5. New measures will be created for each combination of selected measure + function
//
// REQUIREMENTS:
// - Select at least one measure before running the script
// - Functions must have exactly ONE parameter (multi-parameter functions are filtered out)
// - Functions should ideally have the required annotations (see below)
//
// FUNCTION ANNOTATIONS USED:
// The script reads the following annotations from functions to configure the generated measures:
//
// • MeasurePrefix         - Prefix to add to the measure name (e.g., "PY" → "PY Sales")
// • MeasureSuffix         - Suffix to add to the measure name (e.g., "LY" → "Sales LY")
// • FolderPrefix          - Prefix for the display folder path
// • FolderSuffix          - Suffix for the display folder path (e.g., "Last Year")
// • FormatString          - Format string for the new measure (e.g., "#,0.0%")
// • FormatStringExpression - Dynamic format string expression
// • Description           - Description to append to the base measure's description
//
// If MeasurePrefix/MeasureSuffix are missing:  You'll be prompted to enter a suffix for each function
// If FolderPrefix/FolderSuffix are missing: You'll be prompted to enter a folder suffix for each function
// If FormatString is missing: The script uses the base measure's format (or percentage format for 
//                             functions containing "%", "PCT", "IDX", or "INDEX" in their name)
//
// VALIDATION CHECKS:
// - Warns if base measures are missing FormatString/FormatStringExpression
// - Warns if measures or functions are missing descriptions
// - Prevents duplicate measure creation (checks if measure with same expression already exists)
// - Shows a summary of any measures that already exist and were skipped
//
// EXAMPLE:
// Base measure: [Sales] with FormatString "#,0"
// Function: Local_PY with annotation MeasureSuffix = "PY"
// Result: New measure [Sales PY] with expression "Local_PY([Sales])" and format "#,0"
//
// ================================================================================================


using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

// Hide script dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Helper function to count parameters in a function
int CountFunctionParameters(Function function)
{
    string expression = function.Expression;
    
    if (string.IsNullOrWhiteSpace(expression))
        return 0;
    
    int openParenIndex = expression.IndexOf('(');
    int closeParenIndex = -1;
    
    if (openParenIndex == -1)
        return 0;
    
    for (int i = openParenIndex + 1; i < expression.Length - 2; i++)
    {
        if (expression[i] == ')')
        {
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
    
    string parameterSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
    
    // Remove single-line comments before processing
    parameterSection = Regex.Replace(parameterSection, @"//[^\n]*", "");
    
    // Split into individual parameter declarations by comma
    var parameters = parameterSection.Split(',')
        .Select(p => p.Trim())
        .Where(p => !string.IsNullOrWhiteSpace(p))
        .ToList();
    
    // Valid types that can accept a measure
    var validTypes = new[] { "ANYREF", "EXPR" };
    
    // Count only parameters that contain at least one valid type
    int validCount = parameters.Count(p => 
        validTypes.Any(t => Regex.IsMatch(p, @"\b" + t + @"\b", RegexOptions.IgnoreCase))
    );
    
    return validCount;
}

// Always initialize as empty list to prevent null reference errors
List<Function> selectedFunctions = new List<Function>();

// Check if only measures are selected (no functions)
bool onlyMeasuresSelected = Selected.Measures.Any() && !Selected.Functions.Any();

if (onlyMeasuresSelected)
{
    // STEP 1: Get list of functions and filter for TimeIntel prefix and single parameter
    var allFunctions = Model.Functions.ToList();
   // var localFunctions = allFunctions.Where(f => f.Name.StartsWith("Local")).ToList();  //can be excluded, but best practice is starting with Local
    // The above is an example if you only want to include functions with a prefix
    // If so change from all functions to localFunctions below

 
    // Further filter to only include functions with exactly one parameter
    var singleParameterFunctions = allFunctions                                      //replace with allFunctions if not having Local condition
            .Where(f => CountFunctionParameters(f) == 1)
            .OrderBy(f => f) // Sort alphabetically
            .ToList();

    if (allFunctions.Count == 0)
    {
        Error("No functions found in the model.");
        return;
    }

    // Check if there are any single-parameter TimeIntel functions
    if (!singleParameterFunctions.Any())
    {
        Error("No functions with exactly one input parameter found.");
        return;
    }

    // STEP 2: Let user select multiple functions from the filtered list (only single-parameter functions)
    // IMPROVED UI WITH BETTER SPACING AND LAYOUT
    Form functionForm = new Form();
    ListBox functionListBox = new ListBox();
    Button functionOkButton = new Button();
    Button functionCancelButton = new Button();
    Label functionLabel = new Label();

    functionForm.Text = "Select Single-Parameter Functions";
    functionForm.Width = 450;  // Increased width
    functionForm.Height = 580; // Increased height
    functionForm.StartPosition = FormStartPosition.CenterScreen;

    // Improved label with more space
    functionLabel.Text = "Select one or more single-parameter functions:";
    functionLabel.Location = new Point(20, 15);
    functionLabel.Width = 400;  // Increased width
    functionLabel.Height = 30;  // Increased height for multi-line if needed

    functionListBox.Items.AddRange(singleParameterFunctions.Select(f => f.Name).ToArray());
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
    selectedFunctions = singleParameterFunctions.Where(f => selectedFunctionNames.Contains(f.Name)).ToList();
}
else
{
    if (!Selected.Functions.Any())
    {
        Error("No Functions selected.");
        return;
    }
    // Filter selected functions to only include those with one parameter
    selectedFunctions = Selected.Functions.Where(f => CountFunctionParameters(f) == 1).ToList();
    
    if (!selectedFunctions.Any())
    {
        Error("None of the selected functions have exactly one input parameter.");
        return;
    }
}

// Defensive check to prevent null reference error
if (selectedFunctions == null || selectedFunctions.Count == 0)
{
    Error("No single-parameter functions selected.");
    return;
}

// Initialize lists to track missing metadata and already existing measures
var missingFormatString = new List<string>();
var missingMeasureDescription = new List<string>();
var missingFunctionDescription = new List<string>();
var alreadyExistingMeasures = new List<string>();

foreach (var f in selectedFunctions)
{
    if (string.IsNullOrWhiteSpace(f.Description))
        missingFunctionDescription.Add(f.Name);
}

// Check for missing format strings (both FormatString and FormatStringExpression)
foreach (var m in Selected.Measures)
{
    bool hasFormatString = !string.IsNullOrWhiteSpace(m.FormatString);
    bool hasFormatStringExpression = !string.IsNullOrWhiteSpace(m.FormatStringExpression);
    
    if (!hasFormatString && !hasFormatStringExpression)
    {
        missingFormatString.Add(m.Name);
    }

    if (string.IsNullOrWhiteSpace(m.Description))
        missingMeasureDescription.Add(m.Name);
}

// Build info message if any metadata is missing
if (missingFormatString.Any() || missingMeasureDescription.Any() || missingFunctionDescription.Any())
{
    string infoMessage = "Missing metadata detected:\n" +
                     "- Option 1: Press Cancel, fix items, then run script again\n" +
                     "- Option 2: Continue and validate created format string and descriptions\n";
    if (missingFormatString.Any())
    {
        infoMessage += "\n• FormatString or FormatStringExpression missing in measures:\n";
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

// Helper method for input dialog with improved layout
string ShowInputDialog(string text, string caption, string defaultValue = "")
{
    Form prompt = new Form()
    {
        Width = 500,
        Height = 200,
        FormBorderStyle = FormBorderStyle.FixedDialog,
        Text = caption,
        StartPosition = FormStartPosition.CenterScreen,
        MaximizeBox = false,
        MinimizeBox = false
    };
    
    Label textLabel = new Label() 
    { 
        Left = 20, 
        Top = 20, 
        Width = 450, 
        Height = 60, 
        Text = text,
        AutoSize = false
    };
    
    TextBox textBox = new TextBox() 
    { 
        Left = 20, 
        Top = 90, 
        Width = 450,
        Text = defaultValue
    };
    
    Button confirmation = new Button() 
    { 
        Text = "OK", 
        Left = 320, 
        Width = 75, 
        Top = 125, 
        DialogResult = DialogResult.OK 
    };
    
    Button cancel = new Button() 
    { 
        Text = "Cancel", 
        Left = 405, 
        Width = 75, 
        Top = 125, 
        DialogResult = DialogResult.Cancel 
    };
    
    confirmation.Click += (sender, e) => { prompt.Close(); };
    cancel.Click += (sender, e) => { prompt.Close(); };
    
    prompt.Controls.Add(textLabel);
    prompt.Controls.Add(textBox);
    prompt.Controls.Add(confirmation);
    prompt.Controls.Add(cancel);
    prompt.AcceptButton = confirmation;
    prompt.CancelButton = cancel;
    
    return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : string.Empty;
}
// Pre-process functions to get user input for missing suffixes ONCE per function
var functionMeasureSuffixes = new Dictionary<string, string>();

foreach (var f in selectedFunctions)
{
    string measureSuffix = f.GetAnnotation("MeasureSuffix");
    string measurePrefix = f.GetAnnotation("MeasurePrefix");
    
    // Only ask for measure suffix if both prefix and suffix annotations are missing
    if (string. IsNullOrWhiteSpace(measurePrefix) && string.IsNullOrWhiteSpace(measureSuffix))
    {
        string suffix = ShowInputDialog(
            "Enter suffix for the measure name for the selected function:",
            "Define Measure Suffix",
            f. Name  // Default value
        );
        
        functionMeasureSuffixes[f.Name] = suffix ??  "";
    }
}

// Loop through selected measures and functions
foreach (var m in Selected. Measures)
{
    foreach (var f in selectedFunctions)
    {
        string formatString;
        string formatStringExpression;
    
        // First, check if function has a FormatString annotation
        string annotationFormatString = f.GetAnnotation("FormatString");
        string annotationFormatStringExpression = f.GetAnnotation("FormatStringExpression");
        
        if (!string.IsNullOrWhiteSpace(annotationFormatString))
        {
            // Use the annotation value
            formatString = annotationFormatString;
        }
        else
        {
            // Determine the format string based on function name or inherit from measure
            // Check if function name contains percentage indicators
            bool isPercentage = f.Name.Contains("%", StringComparison.OrdinalIgnoreCase) ||
                               f.Name.Contains("pct", StringComparison.OrdinalIgnoreCase) ||
                               f.Name.Contains("idx", StringComparison.OrdinalIgnoreCase) ||
                               f.Name.Contains("index", StringComparison.OrdinalIgnoreCase);
            
            formatString = isPercentage
                            ? "#,0.0%\u003B-#,0.0%\u003B#,0.0%" // Percentage format
                            :  m.FormatString; // Inherit from measure
        }

        if (!string.IsNullOrWhiteSpace(annotationFormatStringExpression))
        {
            // Use the annotation value
            formatStringExpression = annotationFormatStringExpression;
        }
        else
        {        
            formatStringExpression = m.FormatStringExpression; // Inherit from measure
        }

        // Define the expected DAX expression using Function. Name + "(" + m.DaxObjectName + ")"
        string expectedDaxExpression = f.Name + "(" + m. DaxObjectName + ")";

        // Check if a measure with the same expression already exists in the table
        bool measureExists = Model.AllMeasures.Any(existingMeasure =>
            existingMeasure.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") 
            == expectedDaxExpression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "")
            && existingMeasure.Table. Name == m.Table.Name);

        string measurePrefix = f. GetAnnotation("MeasurePrefix");
        string measureSuffix = f.GetAnnotation("MeasureSuffix");

        string newMeasureName;

        // Check if we have prefix or suffix annotations
        if (!string. IsNullOrWhiteSpace(measurePrefix) || !string.IsNullOrWhiteSpace(measureSuffix))
        {
            // Build name with annotations
            string prefix = !string.IsNullOrWhiteSpace(measurePrefix) ? measurePrefix.Trim(' ', '\'') + " " : "";
            string suffix = !string.IsNullOrWhiteSpace(measureSuffix) ? " " + measureSuffix. Trim(' ', '\'') : "";
            
            newMeasureName = (prefix + m.Name + suffix).Trim(' ', '\'');
        }
        else
        {
            // Use the pre-collected suffix for this function
            string suffix = functionMeasureSuffixes[f.Name];
            
            newMeasureName = string.IsNullOrWhiteSpace(suffix) 
                ? m.Name. Trim(' ', '\'')
                : $"{m.Name} {suffix}".Trim(' ', '\'');
        }

        if (measureExists)
        {
            alreadyExistingMeasures.Add($"{newMeasureName} (in {m.Table.Name})");
        }
        else
        {
            string folderPrefix = f.GetAnnotation("FolderPrefix");
            string folderSuffix = f.GetAnnotation("FolderSuffix");

            string displayFolder;

            // Check if we have folder prefix or suffix annotations
            if (!string.IsNullOrWhiteSpace(folderPrefix) || !string.IsNullOrWhiteSpace(folderSuffix))
            {
                // Clean the annotations
                string prefix = !string.IsNullOrWhiteSpace(folderPrefix) ? folderPrefix.Trim() + " " : "";
                string suffix = !string.IsNullOrWhiteSpace(folderSuffix) ? " - " + folderSuffix.Trim() : "";
                
                displayFolder = (prefix + m.DisplayFolder + "\\" + m.Name + suffix).Trim();
            }
            else
            {
                // No annotations, use simple logic
                displayFolder = m.DisplayFolder + "\\" + m.Name;
            }

            // Create the new measure
            var newMeasure = m.Table.AddMeasure(
                newMeasureName, // New measure name
                expectedDaxExpression, // DAX expression
                displayFolder // Display Folder
            );

            // Apply format string and description
            newMeasure.FormatString = formatString;
            newMeasure.FormatStringExpression = formatStringExpression;
            newMeasure.Description = string.IsNullOrWhiteSpace(m.Description) || string.IsNullOrWhiteSpace(f.Description)
                ? ""
                : m.Description + ", " + f.Description;
            
            // Format DAX for readability
            newMeasure. Expression = FormatDax(newMeasure.Expression, shortFormat: true, skipSpaceAfterFunctionName: true);

            // Line shift.
            newMeasure.Expression = $"\n{newMeasure.Expression}";
        }
    }
}
// Show the list of existing measures, if any
if (alreadyExistingMeasures.Any())
{
    string message = "The following measures already exist and were not created:\n\n";
    foreach (var name in alreadyExistingMeasures)
        message += " - " + name + "\n";
    MessageBox.Show(message, "Existing Measures", MessageBoxButtons.OK, MessageBoxIcon.Information);
}