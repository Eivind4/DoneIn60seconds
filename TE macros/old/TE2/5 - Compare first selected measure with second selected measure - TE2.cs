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
#r "System.Drawing"
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Drawing;

// TE2 COMPATIBILITY NOTE - changes made when porting this script from TE3 to TE2:
// 1. Added "#r System.Drawing" (needed explicitly in TE2 for Point/Font; TE3 references it by default).
// 2. Removed duplicate "using System.Linq;" / "using System.Collections.Generic;" (already auto-imported by TE2).
// 3. Converted all $"..." interpolated strings to string concatenation (TE2's compiler doesn't support string interpolation).
// 4. Converted the 3 bare local functions (RemoveComments, CountFunctionParameters, GetFunctionParameterTypes) to
//    Func<> delegate variables, declared in dependency order (RemoveComments first, since the other two call it) -
//    TE2 does not support C# local functions, and variable declarations aren't hoisted like local functions are.
// 5. Replaced the nullable ValueTuple return type "(Measure actual, Measure target)?" with a plain "Measure[]"
//    2-element array (index 0 = actual, index 1 = target) in SelectActualAndTargetMeasures - TE2's older
//    compiler does not support C# 7 ValueTuples.
// No other logic changes were made; behavior should be identical to the TE3 version.

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Helper function to remove comments from parameter section
Func<string, string> RemoveComments = parameterSection =>
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
};

// Improved helper function to count parameters in a function
Func<Function, int> CountFunctionParameters = function =>
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
};

// Returns the declared type (upper-cased, e.g. "ANYREF", "STRING") for each parameter of a
// function, in order. A parameter with no explicit type (e.g. the older untyped "actual, target"
// style) is returned as an empty string - since DAX UDFs without a declared type accept any
// expression (including a measure reference), these are treated the same as ANYREF elsewhere.
Func<Function, List<string>> GetFunctionParameterTypes = function =>
{
    var types = new List<string>();
    string expression = function.Expression;

    if (string.IsNullOrWhiteSpace(expression))
        return types;

    int openParenIndex = expression.IndexOf('(');
    if (openParenIndex == -1)
        return types;

    int closeParenIndex = -1;
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
        return types;

    string parameterSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1).Trim();
    if (string.IsNullOrWhiteSpace(parameterSection))
        return types;

    // Remove comments first, then split on commas (parameter declarations don't contain nested
    // parentheses/commas of their own, so a simple split is sufficient here).
    string cleanParameterSection = RemoveComments(parameterSection);
    var parts = cleanParameterSection.Split(',').Select(p => p.Trim()).Where(p => !string.IsNullOrWhiteSpace(p));

    foreach (var part in parts)
    {
        int colonIndex = part.IndexOf(':');
        types.Add(colonIndex >= 0 ? part.Substring(colonIndex + 1).Trim().ToUpperInvariant() : "");
    }

    return types;
};

// Parameter types that expect a literal/constant value rather than a measure reference, so
// functions declaring any parameter with one of these types cannot be used to build a measure
// from two selected measures and are excluded from the picker below. ANYREF and untyped ("")
// parameters accept a measure reference and remain eligible.
var nonMeasureParameterTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "INT64", "NUMERIC", "STRING", "DECIMAL", "BOOLEAN", "DATETIME", "DOUBLE"
};

// ================================================================================================
// STEP 1: Determine the Actual (measure1) and Target (measure2) measures.
// If exactly two measures are already selected in the model tree, use them directly (as before).
// Otherwise, show a dialog with two search + tree-view pickers (Table > Folder path > Measure)
// to let the user pick the Actual measure first, then the Target measure, in the same dialog.
// ================================================================================================

// (Re)builds a measure-picker TreeView, grouped by Table > DisplayFolder segments > Measure name,
// filtered live by the given search text (matches on measure name or table name). If a
// getCurrentSelection callback is supplied, the leaf matching the currently selected measure is
// re-checked after the rebuild so the selection "sticks" across search-text changes.
Action<TreeView, string, List<Measure>, Func<Measure>> RebuildMeasureTree = (tree, filterText, measures, getCurrentSelection) =>
{
    tree.BeginUpdate();
    tree.Nodes.Clear();

    var currentSelection = getCurrentSelection != null ? getCurrentSelection() : null;

    var filtered = measures
        .Where(m => string.IsNullOrWhiteSpace(filterText) ||
                    m.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Table.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
        .OrderBy(m => m.Table.Name).ThenBy(m => m.DisplayFolder).ThenBy(m => m.Name)
        .ToList();

    var nodeLookup = new Dictionary<string, TreeNode>();

    foreach (var m in filtered)
    {
        var pathParts = new List<string> { m.Table.Name };
        if (!string.IsNullOrWhiteSpace(m.DisplayFolder))
            pathParts.AddRange(m.DisplayFolder.Split('\\').Where(s => !string.IsNullOrWhiteSpace(s)));

        TreeNodeCollection currentLevel = tree.Nodes;
        string pathSoFar = "";
        foreach (var part in pathParts)
        {
            pathSoFar = string.IsNullOrEmpty(pathSoFar) ? part : pathSoFar + "\\" + part;
            TreeNode node;
            if (!nodeLookup.TryGetValue(pathSoFar, out node))
            {
                node = new TreeNode(part);
                currentLevel.Add(node);
                nodeLookup[pathSoFar] = node;
            }
            currentLevel = node.Nodes;
        }

        // Checked is set before the node is added to the tree, so it does not fire AfterCheck.
        var leaf = new TreeNode(m.Name) { Tag = m, Checked = (currentSelection != null && currentSelection == m) };
        currentLevel.Add(leaf);
    }

    tree.ExpandAll();
    tree.EndUpdate();
};

// Shows a single dialog with two independent search + tree-view measure pickers side by side
// ("Actual" on the left, "Target" on the right) and returns the chosen pair as a 2-element array
// (index 0 = actual, index 1 = target), or null if cancelled.
// TE2 note: uses a Measure[] instead of a ValueTuple (TE2's older compiler does not support tuples).
Func<Measure[]> SelectActualAndTargetMeasures = () =>
{
    var allMeasures = Model.AllMeasures.ToList();

    var pickForm = new Form { Text = "Select Actual and Target measures", Width = 1000, Height = 640, StartPosition = FormStartPosition.CenterScreen };

    Measure selectedActual = null;
    Measure selectedTarget = null;

    Func<string, int, int, int, Label> BuildLabel = (text, x, y, w) =>
        new Label { Text = text, Location = new Point(x, y), Width = w, Height = 20, Font = new Font(pickForm.Font, FontStyle.Bold) };

    // ── Actual (left) ──
    var actualLabel     = BuildLabel("Actual measure (measure1):", 20, 15, 440);
    var actualSearch    = new TextBox { Location = new Point(20, 40), Width = 440 };
    var actualTree      = new TreeView { Location = new Point(20, 68), Width = 440, Height = 450, CheckBoxes = true };
    var actualSelLabel  = new Label { Text = "Selected: (none)", Location = new Point(20, 522), Width = 440, Height = 20, Font = new Font(pickForm.Font, FontStyle.Italic) };

    // ── Target (right) ──
    var targetLabel     = BuildLabel("Target measure (measure2):", 500, 15, 440);
    var targetSearch    = new TextBox { Location = new Point(500, 40), Width = 440 };
    var targetTree      = new TreeView { Location = new Point(500, 68), Width = 440, Height = 450, CheckBoxes = true };
    var targetSelLabel  = new Label { Text = "Selected: (none)", Location = new Point(500, 522), Width = 440, Height = 20, Font = new Font(pickForm.Font, FontStyle.Italic) };

    // Wires up a TreeView's checkboxes so that checking a measure leaf node acts as a single-select
    // (any other checked leaf in the same tree is automatically unchecked), while checking a
    // non-leaf (table/folder) node is not a valid selection and is reverted. Also updates the
    // "Selected: <name>" label below the tree so it's obvious what got picked.
    Action<TreeView, Label, Action<Measure>> WireSingleCheckSelection = (tree, selLabel, setSelected) =>
    {
        bool suppress = false;
        tree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;

            var picked = e.Node.Tag as Measure;
            if (picked == null)
            {
                // Folders/tables aren't selectable individually.
                e.Node.Checked = false;
                suppress = false;
                return;
            }

            if (e.Node.Checked)
            {
                // Enforce single selection across the whole tree.
                Action<TreeNodeCollection> UncheckOthers = null;
                UncheckOthers = (nodes) =>
                {
                    foreach (TreeNode n in nodes)
                    {
                        if (n != e.Node && n.Checked) n.Checked = false;
                        UncheckOthers(n.Nodes);
                    }
                };
                UncheckOthers(tree.Nodes);
                setSelected(picked);
                selLabel.Text = "Selected: " + picked.Table.Name + " \\ " + picked.Name;
            }
            else
            {
                setSelected(null);
                selLabel.Text = "Selected: (none)";
            }

            suppress = false;
        };
    };

    WireSingleCheckSelection(actualTree, actualSelLabel, m => selectedActual = m);
    WireSingleCheckSelection(targetTree, targetSelLabel, m => selectedTarget = m);

    RebuildMeasureTree(actualTree, "", allMeasures, () => selectedActual);
    RebuildMeasureTree(targetTree, "", allMeasures, () => selectedTarget);

    actualSearch.TextChanged += (s, e) => RebuildMeasureTree(actualTree, actualSearch.Text, allMeasures, () => selectedActual);
    targetSearch.TextChanged += (s, e) => RebuildMeasureTree(targetTree, targetSearch.Text, allMeasures, () => selectedTarget);

    var okButton     = new Button { Text = "OK",     Location = new Point(380, 560), Width = 80, Height = 30 };
    var cancelButton = new Button { Text = "Cancel", Location = new Point(480, 560), Width = 80, Height = 30, DialogResult = DialogResult.Cancel };

    okButton.Click += (s, e) =>
    {
        if (selectedActual == null || selectedTarget == null)
        {
            MessageBox.Show("Please select both an Actual and a Target measure.", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        pickForm.DialogResult = DialogResult.OK;
        pickForm.Close();
    };

    pickForm.Controls.Add(actualLabel);
    pickForm.Controls.Add(actualSearch);
    pickForm.Controls.Add(actualTree);
    pickForm.Controls.Add(actualSelLabel);
    pickForm.Controls.Add(targetLabel);
    pickForm.Controls.Add(targetSearch);
    pickForm.Controls.Add(targetTree);
    pickForm.Controls.Add(targetSelLabel);
    pickForm.Controls.Add(okButton);
    pickForm.Controls.Add(cancelButton);

    pickForm.AcceptButton = okButton;
    pickForm.CancelButton = cancelButton;

    if (pickForm.ShowDialog() != DialogResult.OK)
        return null;

    return new Measure[] { selectedActual, selectedTarget };
};

// Determine measure1 (Actual) and measure2 (Target)
Measure measure1;
Measure measure2;

if (Selected.Measures.Count == 2)
{
    // Two measures already selected in the model tree - use them directly, as before.
    measure1 = Selected.Measures.ElementAt(0);
    measure2 = Selected.Measures.ElementAt(1);
}
else
{
    var picked = SelectActualAndTargetMeasures();
    if (picked == null)
    {
        Info("Cancelled - no measures were selected.");
        return;
    }
    measure1 = picked[0];
    measure2 = picked[1];
}

List<Function> selectedFunctions = new List<Function>();

// Check if functions are already selected in the model tree (skips the function-picker dialog)
bool functionsPreSelected = Selected.Functions.Any();

if (!functionsPreSelected)
{
    // Adds the 3 standard Comparison DAX UDFs (model-independent) if they don't already exist.
    Action AddMissingComparisonFunctions = () =>
    {
        Action<string, string, string> AddIfMissing = (name, expression, description) =>
        {
            if (Model.Functions.Any(f => f.Name == name)) return;
            var fn = Model.AddFunction(name);
            fn.Expression = expression;
            fn.Description = description;
        };

        AddIfMissing("Comparison.Deviation",
            @"(
    // What you are measuring
    actual,

    // What you want to compare to
    target
) =>
    VAR _Target = target
    VAR _Delta = actual - _Target

    RETURN
        _Delta",
            "Relative difference between measure1 and measure2.");

        AddIfMissing("Comparison.Pct_change",
            @"(
    // What you are measuring
    actual,

    // What you want to compare to
    target
) =>
    VAR _Target = target
    VAR _Delta = actual - _Target
    VAR _Result = DIVIDE ( _Delta, ABS ( _Target ) )
    RETURN
        _Result",
            "Relative percentage difference between measure1 and measure2.");

        AddIfMissing("Comparison.Pct_of_total",
            @"(
    // What you are measuring
    actual,

    // What you want to compare to
    target
) =>
    VAR _Target = target
    VAR _Actual = actual 
    VAR _Result = DIVIDE ( _Actual, _Target ) 
    RETURN
        _Result",
            "Percentage for measure1 of measure2.");
    };

    // STEP 1: Get list of ALL functions with exactly two parameters, excluding any function that
    // declares a parameter with a literal/constant-only type (INT64, NUMERIC, STRING, DECIMAL,
    // BOOLEAN, DATETIME, DOUBLE) since those can't accept a measure reference. Not otherwise
    // restricted by function name, so any measure-compatible 2-parameter DAX UDF (Comparison,
    // ErrorBar, or any other custom one) is selectable here.
    Func<Function, bool> IsTwoParameterFunction = f =>
        CountFunctionParameters(f) == 2 &&
        !GetFunctionParameterTypes(f).Any(t => nonMeasureParameterTypes.Contains(t));

    Func<List<Function>> ComputeTwoParameterComparisonFunctions = () =>
        Model.Functions.Where(IsTwoParameterFunction).ToList();

    var twoParameterComparisonFunctions = ComputeTwoParameterComparisonFunctions();
    var pickedFunctions = new HashSet<Function>();

    // (Re)builds a function-picker TreeView, grouped by dot-separated name segments (e.g.
    // "Comparison.Deviation" -> Comparison > Deviation), filtered live by search text (matches on
    // any part of the function name). Checking a leaf toggles that function in/out of
    // pickedFunctions; checking a group node checks/unchecks all of its descendant leaves.
    Action<TreeView, string, List<Function>> RebuildFunctionTree = (tree, filterText, functions) =>
    {
        tree.BeginUpdate();
        tree.Nodes.Clear();

        var filtered = functions
            .Where(f => string.IsNullOrWhiteSpace(filterText) || f.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
            .OrderBy(f => f.Name)
            .ToList();

        var nodeLookup = new Dictionary<string, TreeNode>();

        foreach (var f in filtered)
        {
            var parts = f.Name.Split('.');
            TreeNodeCollection currentLevel = tree.Nodes;
            string pathSoFar = "";
            for (int i = 0; i < parts.Length - 1; i++)
            {
                pathSoFar = string.IsNullOrEmpty(pathSoFar) ? parts[i] : pathSoFar + "." + parts[i];
                TreeNode node;
                if (!nodeLookup.TryGetValue(pathSoFar, out node))
                {
                    node = new TreeNode(parts[i]);
                    currentLevel.Add(node);
                    nodeLookup[pathSoFar] = node;
                }
                currentLevel = node.Nodes;
            }

            // Checked is set before the node is added to the tree, so it does not fire AfterCheck.
            var leaf = new TreeNode(parts[parts.Length - 1]) { Tag = f, Checked = pickedFunctions.Contains(f) };
            currentLevel.Add(leaf);
        }

        tree.ExpandAll();
        tree.EndUpdate();
    };

    // STEP 2: Let the user select multiple functions via a search box + tree-view (same style as
    // the measure pickers), instead of a flat list.
    Form functionForm = new Form();
    TextBox functionSearchBox = new TextBox();
    TreeView functionTree = new TreeView();
    Button functionOkButton = new Button();
    Button functionMissingButton = new Button();
    Button functionCancelButton = new Button();
    Label functionLabel = new Label();

    functionForm.Text = "Select Two-Parameter Functions";
    functionForm.Width = 450;  // Increased width
    functionForm.Height = 620; // Increased height (search box added)
    functionForm.StartPosition = FormStartPosition.CenterScreen;

    // Improved label with more space
    functionLabel.Text = "Select one or more two-parameter functions (" + twoParameterComparisonFunctions.Count + " available):";
    functionLabel.Location = new Point(20, 15);
    functionLabel.Width = 400;  // Increased width
    functionLabel.Height = 30;  // Increased height for multi-line if needed

    functionSearchBox.Location = new Point(20, 48);
    functionSearchBox.Width = 400;

    functionTree.CheckBoxes = true;
    functionTree.Location = new Point(20, 76);
    functionTree.Width = 400;  // Increased width
    functionTree.Height = 404; // Increased height

    bool suppressFunctionCheckEvents = false;
    functionTree.AfterCheck += (s, e) =>
    {
        if (suppressFunctionCheckEvents) return;
        suppressFunctionCheckEvents = true;

        var fn = e.Node.Tag as Function;
        if (fn != null)
        {
            if (e.Node.Checked) pickedFunctions.Add(fn);
            else pickedFunctions.Remove(fn);
        }
        else
        {
            // Group node - propagate the check state down to all descendant leaves.
            Action<TreeNodeCollection, bool> SetAll = null;
            SetAll = (nodes, val) =>
            {
                foreach (TreeNode n in nodes)
                {
                    n.Checked = val;
                    var leafFn = n.Tag as Function;
                    if (leafFn != null)
                    {
                        if (val) pickedFunctions.Add(leafFn);
                        else pickedFunctions.Remove(leafFn);
                    }
                    SetAll(n.Nodes, val);
                }
            };
            SetAll(e.Node.Nodes, e.Node.Checked);
        }

        suppressFunctionCheckEvents = false;
    };

    RebuildFunctionTree(functionTree, "", twoParameterComparisonFunctions);
    functionSearchBox.TextChanged += (s, e) => RebuildFunctionTree(functionTree, functionSearchBox.Text, twoParameterComparisonFunctions);

    // Centered and properly spaced buttons - OK, "Add Comparison DaxUDF" (creates the standard
    // Comparison functions and refreshes the tree above without closing the dialog), and Cancel.
    int buttonY = 540;  // Lower position due to increased form height
    int okWidth = 80, missingWidth = 190, cancelWidth = 80, buttonSpacing = 15, buttonHeight = 30;
    int totalButtonWidth = okWidth + missingWidth + cancelWidth + (buttonSpacing * 2);
    int startX = (functionForm.Width - totalButtonWidth) / 2;

    functionOkButton.Text = "OK";
    functionOkButton.Location = new Point(startX, buttonY);
    functionOkButton.Width = okWidth;
    functionOkButton.Height = buttonHeight;
    // Validated manually below (instead of a plain DialogResult) so an empty selection doesn't
    // close the dialog.

    functionMissingButton.Text = "Add Comparison DaxUDF";
    functionMissingButton.Location = new Point(startX + okWidth + buttonSpacing, buttonY);
    functionMissingButton.Width = missingWidth;
    functionMissingButton.Height = buttonHeight;
    // No DialogResult - creates the functions, refreshes the tree, and keeps the dialog open.

    functionCancelButton.Text = "Cancel";
    functionCancelButton.Location = new Point(startX + okWidth + missingWidth + (buttonSpacing * 2), buttonY);
    functionCancelButton.Width = cancelWidth;
    functionCancelButton.Height = buttonHeight;
    functionCancelButton.DialogResult = DialogResult.Cancel;

    functionOkButton.Click += (s, e) =>
    {
        if (!pickedFunctions.Any())
        {
            MessageBox.Show("Please select at least one function, or click 'Add Comparison DaxUDF' to create the standard ones first.", "No function selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        functionForm.DialogResult = DialogResult.OK;
        functionForm.Close();
    };

    functionMissingButton.Click += (s, e) =>
    {
        AddMissingComparisonFunctions();

        twoParameterComparisonFunctions = ComputeTwoParameterComparisonFunctions();
        RebuildFunctionTree(functionTree, functionSearchBox.Text, twoParameterComparisonFunctions);

        functionLabel.Text = "Select one or more two-parameter functions (" + twoParameterComparisonFunctions.Count + " available):";
    };

    functionForm.Controls.Add(functionLabel);
    functionForm.Controls.Add(functionSearchBox);
    functionForm.Controls.Add(functionTree);
    functionForm.Controls.Add(functionOkButton);
    functionForm.Controls.Add(functionMissingButton);
    functionForm.Controls.Add(functionCancelButton);

    functionForm.AcceptButton = functionOkButton;
    functionForm.CancelButton = functionCancelButton;

    var dialogResult = functionForm.ShowDialog();

    if (dialogResult == DialogResult.Cancel)
    {
        return;
    }

    selectedFunctions = pickedFunctions.ToList();
}
else
{
    // Filter selected functions to only include measure-compatible functions with exactly two
    // parameters (no restriction on function name).
    selectedFunctions = Selected.Functions.Where(f =>
        CountFunctionParameters(f) == 2 &&
        !GetFunctionParameterTypes(f).Any(t => nonMeasureParameterTypes.Contains(t))).ToList();

    if (!selectedFunctions.Any())
    {
        Error("None of the selected functions have exactly two measure-compatible input parameters.");
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

// Check the Actual/Target measures (not Selected.Measures - these may have been picked via the
// tree-view dialog above rather than the model tree selection)
foreach (var m in new[] { measure1, measure2 })
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

// Normalizes a DAX expression for comparison purposes (collapses all whitespace), so that
// formatting differences (spacing, line breaks) don't cause a false "different expression" result.
Func<string, string> NormalizeDax = s => string.Join(" ", (s ?? "").Split(
    new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));

// Shown when a new measure's name collides with an existing measure that has a different
// expression. Lets the user Skip creating the new measure, or Rename it (pre-filled with a
// comparison-oriented suggestion). Returns the chosen name, or null if the user chose to skip.
Func<string, string, string> ResolveMeasureNameConflict = (originalName, suggestedName) =>
{
    var conflictForm = new Form
    {
        Text = "Measure name already exists",
        Width = 480,
        Height = 230,
        FormBorderStyle = FormBorderStyle.FixedDialog,
        StartPosition = FormStartPosition.CenterScreen,
        MaximizeBox = false,
        MinimizeBox = false
    };

    var infoLabel = new Label
    {
        Text = "A measure named '" + originalName + "' already exists with a different expression.\nChoose how to proceed:",
        Location = new Point(15, 15),
        Width = 440,
        Height = 40
    };

    var renameRadio = new RadioButton { Text = "Rename new measure to:", Location = new Point(15, 65), Width = 160, Checked = true };
    var renameBox   = new TextBox { Location = new Point(180, 63), Width = 275, Text = suggestedName };
    var skipRadio   = new RadioButton { Text = "Skip - do not create this measure", Location = new Point(15, 100), Width = 400 };

    var okButton     = new Button { Text = "OK", Location = new Point(285, 150), Width = 80, Height = 30 };
    var cancelButton = new Button { Text = "Cancel (skip)", Location = new Point(375, 150), Width = 90, Height = 30 };

    renameRadio.CheckedChanged += (s, e) => renameBox.Enabled = renameRadio.Checked;

    string result = null;
    okButton.Click += (s, e) =>
    {
        if (renameRadio.Checked && string.IsNullOrWhiteSpace(renameBox.Text))
        {
            MessageBox.Show("Please enter a name, or choose 'Skip'.", "Name required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        result = skipRadio.Checked ? null : renameBox.Text.Trim();
        conflictForm.DialogResult = DialogResult.OK;
        conflictForm.Close();
    };
    cancelButton.Click += (s, e) =>
    {
        result = null;
        conflictForm.DialogResult = DialogResult.OK;
        conflictForm.Close();
    };

    conflictForm.Controls.Add(infoLabel);
    conflictForm.Controls.Add(renameRadio);
    conflictForm.Controls.Add(renameBox);
    conflictForm.Controls.Add(skipRadio);
    conflictForm.Controls.Add(okButton);
    conflictForm.Controls.Add(cancelButton);
    conflictForm.AcceptButton = okButton;

    conflictForm.ShowDialog();

    return result;
};

foreach (var function in selectedFunctions)
{
    // Generate measure name based on function name patterns
    string measureName;
    string functionNameLower = function.Name.ToLower();
    
    if (functionNameLower.Contains("errorbar"))
    {
        if (functionNameLower.Contains("above"))
            measureName = measure1.Name + " EB Above";
        else if (functionNameLower.Contains("below"))
            measureName = measure1.Name + " EB Below";
        else
            measureName = measure1.Name + " EB";
    }
    else if (function.Name.EndsWith("Deviation"))
    {
        measureName = measure1.Name + " Dev";
    }
    else if (functionNameLower.Contains("pct_of_total"))
    {
        measureName = measure1.Name + " % of target";
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

    // Construct DAX expression: f.Name(measure1, measure2)
    var daxExpression = function.Name + "(" + measure1.DaxObjectName + ", " + measure2.DaxObjectName + ")";

    // If a measure with the exact same expression already exists anywhere in the target table,
    // there's nothing new to create - skip silently.
    var existingByExpression = finalTargetTable.Measures.FirstOrDefault(
        m => NormalizeDax(m.Expression) == NormalizeDax(daxExpression));
    if (existingByExpression != null)
    {
        alreadyExistingMeasures.Add(measureName + " (identical measure already exists as '" + existingByExpression.Name + "')");
        continue;
    }

    // If a measure with this exact NAME already exists but with a different expression, ask the
    // user whether to skip it or provide an alternate name (pre-filled with a comparison-oriented
    // suggestion) instead of silently overwriting or colliding.
    var existingByName = finalTargetTable.Measures.FirstOrDefault(m => m.Name == measureName);
    if (existingByName != null)
    {
        string suggestedName = measureName + " (" + measure1.Name + " vs " + measure2.Name + ")";
        string resolvedName = ResolveMeasureNameConflict(measureName, suggestedName);

        if (resolvedName == null)
        {
            alreadyExistingMeasures.Add(measureName + " (skipped - name already exists with a different expression)");
            continue;
        }

        if (finalTargetTable.Measures.Any(m => m.Name == resolvedName))
        {
            alreadyExistingMeasures.Add(resolvedName + " (skipped - the alternate name also already exists)");
            continue;
        }

        measureName = resolvedName;
    }

    // Create the new measure - placed in a new sub-folder under measure1's existing DisplayFolder
    // (or just the sub-folder name if measure1 has no folder): "Error bars" for ErrorBar-based
    // functions, "Comparison" for all others - so it's a new folder branching off from where
    // measure1 lives.
    string comparisonSubFolder = functionNameLower.Contains("errorbar") ? "Error bars" : "Comparison";
    string comparisonFolder = string.IsNullOrWhiteSpace(measure1.DisplayFolder)
        ? comparisonSubFolder
        : measure1.DisplayFolder + "\\" + comparisonSubFolder;

    var newMeasure = finalTargetTable.AddMeasure(
        measureName,
        daxExpression,
        comparisonFolder
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
    string description = function.Description ?? ("Comparison function " + function.Name);
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