#r "System.Drawing"

// ================================================================================================
// Title: Create Measures from model dependent DaxUDF
// Author: Eivind Haugen
// ================================================================================================
//
// TE2 COMPATIBILITY NOTE:
// This is a TE2-adapted copy of the TE3 version of this script. TE2's default (non-Roslyn) script
// compiler only supports older C# syntax and requires explicit assembly references, so the
// following adjustments were made:
// • Added "#r \"System.Drawing\"" at the top - System.Drawing (Point/Size/Font/etc.) is not
//   referenced by default in TE2, unlike TE3
// • String interpolation ($"...") -> replaced with plain string concatenation
// • Local functions (e.g. "int Foo(...) { }") -> rewritten as Func<>/Action<> lambda variables
// • Tuple return type ((string, string, string)) -> replaced with Dictionary<string,string>
//   (keys "Prefix"/"Suffix"/"FolderName")
// • Type pattern matching ("is Function taggedFunction") -> replaced with classic "is" check + cast
// • string.Contains(x, StringComparison) (a newer BCL overload) -> replaced with
//   string.IndexOf(x, StringComparison) >= 0
// • Removed redundant "using System.Linq;"/"using System.Collections.Generic;" (already
//   auto-imported by TE2, causing CS0105 warnings)
// If your TE2 installation has a Roslyn compiler configured (File > Preferences), the original
// TE3 syntax would also work, but this version keeps compatibility with the default compiler.
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
// If MeasurePrefix/MeasureSuffix are missing: You'll be prompted (once per function) with a dialog
//    where you can enter a Prefix and/or a Suffix (at least one required). A live preview shows how
//    the resulting measure name would look, based on the first selected measure as an example.
// If FolderPrefix/FolderSuffix are missing: The same dialog also offers an OPTIONAL folder name to
//    append to the measure's existing display folder (e.g. DisplayFolder\FolderName). A live preview
//    shows the resulting folder path. Leave blank to keep the current fallback (DisplayFolder\).
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

// Hide script dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Helper function to count parameters in a function
// (TE2 compatibility: written as a Func<> lambda instead of a local function, since the
// default TE2 script compiler does not support C# 7+ local function declarations)
Func<Function, int> CountFunctionParameters = (function) =>
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
};

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
    // UI: a namespace-hierarchy TreeView (functions are grouped by the dot-separated segments of their
    // name, e.g. "Local.TimeIntel.TI_YTD" -> Local > TimeIntel > TI_YTD) with checkboxes, so the user can
    // expand/collapse namespaces and check individual functions OR an entire namespace node at once
    // (checking a namespace node cascades the check down to every function beneath it). A search box
    // above the tree filters the list live to only names containing the typed text (case-insensitive),
    // while preserving any selections already made.
    Form functionForm = new Form();
    Label searchLabel = new Label();
    TextBox searchBox = new TextBox();
    TreeView functionTree = new TreeView();
    Button functionOkButton = new Button();
    Button functionCancelButton = new Button();
    Label functionLabel = new Label();

    functionForm.Text = "Select Single-Parameter Functions";
    functionForm.Width = 500;
    functionForm.Height = 650;
    functionForm.StartPosition = FormStartPosition.CenterScreen;

    functionLabel.Text = "Select one or more single-parameter functions (expand namespaces, or check a namespace to select all functions under it):";
    functionLabel.Location = new Point(20, 15);
    functionLabel.Width = 440;
    functionLabel.Height = 45;

    searchLabel.Text = "Search:";
    searchLabel.Location = new Point(20, 65);
    searchLabel.AutoSize = true;

    searchBox.Location = new Point(75, 62);
    searchBox.Width = 385;

    functionTree.Location = new Point(20, 90);
    functionTree.Width = 440;
    functionTree.Height = 470;
    functionTree.CheckBoxes = true;

    // Full function name -> checked state, kept across search-filter rebuilds of the tree
    var checkedFunctionNames = new HashSet<string>();
    bool suppressTreeEvents = false;

    Action<TreeNode> UpdateAncestors = null;
    UpdateAncestors = (node) =>
    {
        var parent = node.Parent;
        while (parent != null)
        {
            parent.Checked = parent.Nodes.Cast<TreeNode>().All(n => n.Checked);
            parent = parent.Parent;
        }
    };

    Action<TreeNode, bool> SetDescendantsChecked = null;
    SetDescendantsChecked = (node, isChecked) =>
    {
        foreach (TreeNode child in node.Nodes)
        {
            child.Checked = isChecked;
            SetDescendantsChecked(child, isChecked);
        }
    };

    Action<TreeNode> SyncLeafSelections = null;
    SyncLeafSelections = (node) =>
    {
        if (node.Tag != null && node.Tag is Function)
        {
            var taggedFunction = (Function)node.Tag;
            if (node.Checked) checkedFunctionNames.Add(taggedFunction.Name);
            else checkedFunctionNames.Remove(taggedFunction.Name);
        }
        foreach (TreeNode child in node.Nodes)
            SyncLeafSelections(child);
    };

    // (Re)builds the tree from only the functions matching the current search text, restoring
    // previously checked functions (by name) so filtering never loses a selection.
    Action<string> RebuildFunctionTree = (filterText) =>
    {
        suppressTreeEvents = true;
        functionTree.BeginUpdate();
        functionTree.Nodes.Clear();

        var filtered = singleParameterFunctions
            .Where(f => string.IsNullOrWhiteSpace(filterText) ||
                        f.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
            .OrderBy(f => f.Name)
            .ToList();

        var nodeLookup = new Dictionary<string, TreeNode>();

        foreach (var f in filtered)
        {
            var segments = f.Name.Split('.');
            TreeNodeCollection currentLevel = functionTree.Nodes;
            string pathSoFar = "";

            for (int p = 0; p < segments.Length; p++)
            {
                pathSoFar = p == 0 ? segments[p] : pathSoFar + "." + segments[p];
                bool isLeaf = p == segments.Length - 1;

                TreeNode node;
                if (!nodeLookup.TryGetValue(pathSoFar, out node))
                {
                    node = new TreeNode(segments[p]);
                    currentLevel.Add(node);
                    nodeLookup[pathSoFar] = node;
                }

                if (isLeaf)
                {
                    node.Tag = f;
                    node.Checked = checkedFunctionNames.Contains(f.Name);
                }

                currentLevel = node.Nodes;
            }
        }

        // Reflect children state on namespace (non-leaf) nodes, bottom-up
        Action<TreeNode> RollUpCheckedState = null;
        RollUpCheckedState = (node) =>
        {
            foreach (TreeNode child in node.Nodes)
                RollUpCheckedState(child);

            if (node.Tag == null && node.Nodes.Count > 0)
                node.Checked = node.Nodes.Cast<TreeNode>().All(n => n.Checked);
        };
        foreach (TreeNode root in functionTree.Nodes)
            RollUpCheckedState(root);

        functionTree.ExpandAll();
        functionTree.EndUpdate();
        suppressTreeEvents = false;
    };

    RebuildFunctionTree("");

    searchBox.TextChanged += (s, e) => RebuildFunctionTree(searchBox.Text);

    functionTree.AfterCheck += (s, e) =>
    {
        if (suppressTreeEvents) return;
        suppressTreeEvents = true;

        SetDescendantsChecked(e.Node, e.Node.Checked);
        SyncLeafSelections(e.Node);
        UpdateAncestors(e.Node);

        suppressTreeEvents = false;
    };

    // Centered and properly spaced buttons
    int buttonY = 570;
    int buttonWidth = 80;
    int buttonSpacing = 20;
    int totalButtonWidth = (buttonWidth * 2) + buttonSpacing;
    int startX = (functionForm.Width - totalButtonWidth) / 2;

    functionOkButton.Text = "OK";
    functionOkButton.Location = new Point(startX, buttonY);
    functionOkButton.Width = buttonWidth;
    functionOkButton.Height = 30;
    functionOkButton.DialogResult = DialogResult.OK;

    functionCancelButton.Text = "Cancel";
    functionCancelButton.Location = new Point(startX + buttonWidth + buttonSpacing, buttonY);
    functionCancelButton.Width = buttonWidth;
    functionCancelButton.Height = 30;
    functionCancelButton.DialogResult = DialogResult.Cancel;

    functionForm.Controls.Add(functionLabel);
    functionForm.Controls.Add(searchLabel);
    functionForm.Controls.Add(searchBox);
    functionForm.Controls.Add(functionTree);
    functionForm.Controls.Add(functionOkButton);
    functionForm.Controls.Add(functionCancelButton);

    functionForm.AcceptButton = functionOkButton;
    functionForm.CancelButton = functionCancelButton;

    var dialogResult = functionForm.ShowDialog();

    if (dialogResult == DialogResult.Cancel)
    {
        return;
    }

    selectedFunctions = singleParameterFunctions.Where(f => checkedFunctionNames.Contains(f.Name)).ToList();
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

// Helper: prompts for the measure naming (Prefix and/or Suffix - at least one required, since
// Cancel/blank on both is not a valid name), and/or an optional Folder name to append - showing a
// live example based on a sample measure (the first selected measure) as the user types. Only the
// section(s) actually needed (needMeasureNaming / needFolderNaming) are shown.
// TE2 compatibility: returns a Dictionary<string,string> (keys "Prefix"/"Suffix"/"FolderName")
// instead of a ValueTuple, since the default TE2 script compiler does not support C# 7+ tuples.
Func<Function, bool, bool, Measure, Dictionary<string, string>> ShowFunctionNamingDialog = (f, needMeasureNaming, needFolderNaming, sampleMeasure) =>
{
    var form = new Form
    {
        Text = "Define naming for function '" + f.Name + "'",
        Width = 560,
        FormBorderStyle = FormBorderStyle.FixedDialog,
        StartPosition = FormStartPosition.CenterScreen,
        MaximizeBox = false,
        MinimizeBox = false,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink
    };
    var flow = new FlowLayoutPanel { Padding = new Padding(15), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    form.Controls.Add(flow);

    string sampleMeasureName = sampleMeasure != null ? sampleMeasure.Name : "<Measure>";
    string sampleFolder      = sampleMeasure != null ? sampleMeasure.DisplayFolder : "";

    TextBox prefixBox = null, suffixBox = null, folderBox = null;
    Label measurePreview = null, folderPreview = null;

    if (needMeasureNaming)
    {
        flow.Controls.Add(new Label { Text = "Measure name - enter a Prefix and/or Suffix (at least one required):", AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 0, 0, 8) });

        var prefixRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        prefixRow.Controls.Add(new Label { Text = "Prefix:", Width = 60, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 0, 0) });
        prefixBox = new TextBox { Width = 400 };
        prefixRow.Controls.Add(prefixBox);
        flow.Controls.Add(prefixRow);

        var suffixRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        suffixRow.Controls.Add(new Label { Text = "Suffix:", Width = 60, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 0, 0) });
        suffixBox = new TextBox { Width = 400, Text = f.Name };
        suffixRow.Controls.Add(suffixBox);
        flow.Controls.Add(suffixRow);

        measurePreview = new Label { AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 2, 0, 10), Font = new Font(form.Font, FontStyle.Italic) };
        flow.Controls.Add(measurePreview);
    }

    if (needFolderNaming)
    {
        flow.Controls.Add(new Label { Text = "Folder name to append (optional):", AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 6, 0, 8) });
        folderBox = new TextBox { Width = 466 };
        flow.Controls.Add(folderBox);
        folderPreview = new Label { AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 4, 0, 10), Font = new Font(form.Font, FontStyle.Italic) };
        flow.Controls.Add(folderPreview);
    }

    Action updateMeasurePreview = () =>
    {
        if (measurePreview == null) return;
        string prefixTxt = !string.IsNullOrWhiteSpace(prefixBox.Text) ? prefixBox.Text.Trim(' ', '\'') + " " : "";
        string suffixTxt = !string.IsNullOrWhiteSpace(suffixBox.Text) ? " " + suffixBox.Text.Trim(' ', '\'') : "";
        string example = (prefixTxt + sampleMeasureName + suffixTxt).Trim(' ', '\'');
        measurePreview.Text = "Example (for '" + sampleMeasureName + "'): " + example;
    };
    Action updateFolderPreview = () =>
    {
        if (folderPreview == null) return;
        string folderTxt = folderBox.Text.Trim();
        string example = string.IsNullOrWhiteSpace(folderTxt)
            ? sampleFolder + "\\"
            : (sampleFolder + "\\" + folderTxt).Trim('\\');
        folderPreview.Text = "Example folder (for '" + sampleMeasureName + "'): " + example;
    };

    if (needMeasureNaming)
    {
        prefixBox.TextChanged += (s, e) => updateMeasurePreview();
        suffixBox.TextChanged += (s, e) => updateMeasurePreview();
        updateMeasurePreview();
    }
    if (needFolderNaming)
    {
        folderBox.TextChanged += (s, e) => updateFolderPreview();
        updateFolderPreview();
    }

    var btnPanel  = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
    var okBtn     = new Button { Text = "OK", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    var cancelBtn = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    btnPanel.Controls.Add(okBtn);
    btnPanel.Controls.Add(cancelBtn);
    flow.Controls.Add(btnPanel);
    form.CancelButton = cancelBtn;

    okBtn.Click += (s, e) =>
    {
        if (needMeasureNaming && string.IsNullOrWhiteSpace(prefixBox.Text) && string.IsNullOrWhiteSpace(suffixBox.Text))
        {
            MessageBox.Show("Enter at least a Prefix or a Suffix for the measure name.", "Naming required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        form.DialogResult = DialogResult.OK;
        form.Close();
    };
    form.AcceptButton = okBtn;

    var dialogResult = form.ShowDialog();
    if (dialogResult != DialogResult.OK)
        return new Dictionary<string, string> { { "Prefix", "" }, { "Suffix", "" }, { "FolderName", "" } };

    return new Dictionary<string, string>
    {
        { "Prefix", prefixBox != null ? prefixBox.Text.Trim() : "" },
        { "Suffix", suffixBox != null ? suffixBox.Text.Trim() : "" },
        { "FolderName", folderBox != null ? folderBox.Text.Trim() : "" }
    };
};

// Pre-process functions to get user input for missing measure Prefix/Suffix and/or Folder name,
// ONCE per function. The live preview inside the dialog uses the first selected measure (not
// required for every measure - it is just an illustrative example).
var functionMeasurePrefixes = new Dictionary<string, string>();
var functionMeasureSuffixes = new Dictionary<string, string>();
var functionFolderNames     = new Dictionary<string, string>();
var sampleMeasureForPreview = Selected.Measures.FirstOrDefault();

foreach (var f in selectedFunctions)
{
    string measureSuffix = f.GetAnnotation("MeasureSuffix");
    string measurePrefix = f.GetAnnotation("MeasurePrefix");
    bool needsMeasureNaming = string.IsNullOrWhiteSpace(measurePrefix) && string.IsNullOrWhiteSpace(measureSuffix);

    string folderPrefixAnn = f.GetAnnotation("FolderPrefix");
    string folderSuffixAnn = f.GetAnnotation("FolderSuffix");
    bool needsFolderNaming = string.IsNullOrWhiteSpace(folderPrefixAnn) && string.IsNullOrWhiteSpace(folderSuffixAnn);

    if (needsMeasureNaming || needsFolderNaming)
    {
        var naming = ShowFunctionNamingDialog(f, needsMeasureNaming, needsFolderNaming, sampleMeasureForPreview);

        if (needsMeasureNaming)
        {
            functionMeasurePrefixes[f.Name] = naming["Prefix"];
            functionMeasureSuffixes[f.Name] = naming["Suffix"];
        }
        if (needsFolderNaming)
        {
            functionFolderNames[f.Name] = naming["FolderName"];
        }
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
            // TE2 compatibility: the .NET Framework String class does not have the newer
            // Contains(string, StringComparison) overload, so IndexOf(...) >= 0 is used instead
            bool isPercentage = f.Name.IndexOf("%", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               f.Name.IndexOf("pct", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               f.Name.IndexOf("idx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               f.Name.IndexOf("index", StringComparison.OrdinalIgnoreCase) >= 0;
            
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

        string measurePrefix = f.GetAnnotation("MeasurePrefix");
        string measureSuffix = f.GetAnnotation("MeasureSuffix");

        // If neither annotation exists, fall back to the Prefix/Suffix the user entered once
        // for this function in the naming dialog (functionMeasurePrefixes/functionMeasureSuffixes)
        if (string.IsNullOrWhiteSpace(measurePrefix) && string.IsNullOrWhiteSpace(measureSuffix))
        {
            measurePrefix = functionMeasurePrefixes.ContainsKey(f.Name) ? functionMeasurePrefixes[f.Name] : "";
            measureSuffix = functionMeasureSuffixes.ContainsKey(f.Name) ? functionMeasureSuffixes[f.Name] : "";
        }

        string prefixPart = !string.IsNullOrWhiteSpace(measurePrefix) ? measurePrefix.Trim(' ', '\'') + " " : "";
        string suffixPart = !string.IsNullOrWhiteSpace(measureSuffix) ? " " + measureSuffix.Trim(' ', '\'') : "";
        string newMeasureName = (prefixPart + m.Name + suffixPart).Trim(' ', '\'');

        if (measureExists)
        {
            alreadyExistingMeasures.Add(newMeasureName + " (in " + m.Table.Name + ")");
        }
        else
        {
            string folderPrefix = f.GetAnnotation("FolderPrefix");
            string folderSuffix = f.GetAnnotation("FolderSuffix");

            // If neither annotation exists, fall back to the (optional) folder name the user
            // entered once for this function in the naming dialog (functionFolderNames)
            if (string.IsNullOrWhiteSpace(folderPrefix) && string.IsNullOrWhiteSpace(folderSuffix)
                && functionFolderNames.ContainsKey(f.Name))
            {
                folderSuffix = functionFolderNames[f.Name];
            }

            string displayFolder;

            // Check if we have folder prefix or suffix annotations (or a user-entered folder name)
            if (!string.IsNullOrWhiteSpace(folderPrefix) || !string.IsNullOrWhiteSpace(folderSuffix))
            {
                // Clean the annotations
                string prefix = !string.IsNullOrWhiteSpace(folderPrefix) ? folderPrefix.Trim() + " " : "";
                string suffix = !string.IsNullOrWhiteSpace(folderSuffix) ? folderSuffix.Trim() : "";
                
                displayFolder = (prefix + m.DisplayFolder + "\\" + suffix).Trim();
            }
            else
            {
                // No annotations and no folder name entered, use simple logic
                displayFolder = m.DisplayFolder + "\\";
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
            
            // NOTE: FormatDax() call removed for TE2 - not needed since these are simple DAX
            // expressions (a single function call), and repeated calls trigger the DaxFormatter.com
            // throttling warning. Add it back manually if you switch to more complex expressions.

            // Line shift.
            newMeasure.Expression = "\n" + newMeasure.Expression;
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