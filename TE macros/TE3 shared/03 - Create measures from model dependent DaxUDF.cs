// ================================================================================================
// Title: Create Measures from model dependent DaxUDF
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
// - Detects base measures missing FormatString/FormatStringExpression, and measures or functions
//   missing descriptions. They are listed in an editable "Missing Metadata" grid where you can:
//     * Apply changes and continue - fill in format string OR format string expression (not both)
//       and/or descriptions; values are saved on the existing measures / DAX UDFs and inherited by
//       the new measures
//     * Continue without changes, or Cancel
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
    // UI: a namespace-hierarchy TreeView (functions are grouped by the dot-separated segments of their
    // name, e.g. "Local.TimeIntel.TI_YTD" -> Local > TimeIntel > TI_YTD) with checkboxes, so the user can
    // expand/collapse namespaces and check individual functions OR an entire namespace node at once
    // (checking a namespace node cascades the check down to every function beneath it). A search box
    // above the tree filters the list live to only names containing the typed text (case-insensitive),
    // while preserving any selections already made.
    // Responsive layout: the dialog is resizable, the tree fills the free space and the
    // buttons stay visible at the bottom at any size / display scaling.
    var uiFont = new Font("Segoe UI", 9);
    var workingArea = Screen.PrimaryScreen.WorkingArea;
    Form functionForm = new Form {
        Text = "Select Single-Parameter Functions",
        Size = new Size(Math.Min(720, (int)(workingArea.Width * 0.6)), (int)(workingArea.Height * 0.8)),
        MinimumSize = new Size(480, 420),
        StartPosition = FormStartPosition.CenterScreen,
        FormBorderStyle = FormBorderStyle.Sizable,
        MaximizeBox = true,
        MinimizeBox = false,
        Font = uiFont
    };
    var functionLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 5, Padding = new Padding(12) };
    functionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    functionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    functionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    functionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    functionLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    functionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    functionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    functionForm.Controls.Add(functionLayout);

    Label functionLabel = new Label {
        Text = "Select one or more single-parameter functions (expand namespaces, or check a namespace to select all functions under it):",
        AutoSize = true, MaximumSize = new Size(900, 0), Margin = new Padding(0, 0, 0, 8),
        Font = new Font("Segoe UI", 10, FontStyle.Bold)
    };
    Label searchLabel = new Label { Text = "Search:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 8) };
    TextBox searchBox = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 8) };
    TreeView functionTree = new TreeView { Dock = DockStyle.Fill, CheckBoxes = true, HideSelection = false };
    Label functionStatus = new Label { Text = "0 function(s) selected", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
    Button functionOkButton = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    Button functionCancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    var functionButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    functionButtons.Controls.Add(functionOkButton);
    functionButtons.Controls.Add(functionCancelButton);

    functionLayout.Controls.Add(functionLabel, 0, 0);
    functionLayout.SetColumnSpan(functionLabel, 2);
    functionLayout.Controls.Add(searchLabel, 0, 1);
    functionLayout.Controls.Add(searchBox, 1, 1);
    functionLayout.Controls.Add(functionTree, 0, 2);
    functionLayout.SetColumnSpan(functionTree, 2);
    functionLayout.Controls.Add(functionStatus, 0, 3);
    functionLayout.SetColumnSpan(functionStatus, 2);
    functionLayout.Controls.Add(functionButtons, 0, 4);
    functionLayout.SetColumnSpan(functionButtons, 2);

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
        if (node.Tag is Function taggedFunction)
        {
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
        functionStatus.Text = checkedFunctionNames.Count + " function(s) selected";
    };

    functionOkButton.Click += (s, e) =>
    {
        if (checkedFunctionNames.Count == 0)
        {
            MessageBox.Show("Select at least one function.", "No functions selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        functionForm.DialogResult = DialogResult.OK;
        functionForm.Close();
    };

    functionForm.AcceptButton = functionOkButton;
    functionForm.CancelButton = functionCancelButton;
    functionForm.Shown += (s, e) => searchBox.Focus();

    var dialogResult = functionForm.ShowDialog();

    if (dialogResult != DialogResult.OK)
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

// Missing metadata: one editable grid instead of a message box. Fill in the highlighted cells
// to fix the base measures / DAX UDFs directly in the model before the new measures are created;
// the new measures then inherit the values (format from the base measure, description =
// measure description + ", " + function description). Blank cells stay as they are.
if (missingFormatString.Any() || missingMeasureDescription.Any() || missingFunctionDescription.Any())
{
    var editBack       = Color.FromArgb(255, 244, 179);   // highlighted editable cells
    var editHeaderBack = Color.FromArgb(255, 214, 102);
    var naBack         = Color.FromArgb(235, 235, 235);   // not applicable for this row
    var boldFont       = new Font("Segoe UI", 9, FontStyle.Bold);
    var workArea       = Screen.PrimaryScreen.WorkingArea;

    var metadataForm = new Form {
        Text = "Missing Metadata",
        Size = new Size((int)(workArea.Width * 0.75), (int)(workArea.Height * 0.6)),
        MinimumSize = new Size(760, 380),
        StartPosition = FormStartPosition.CenterScreen,
        FormBorderStyle = FormBorderStyle.Sizable,
        MaximizeBox = true,
        MinimizeBox = false,
        Font = new Font("Segoe UI", 9)
    };
    var metadataLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(12) };
    metadataLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    metadataLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    metadataLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    metadataLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    metadataForm.Controls.Add(metadataLayout);

    var metadataHeading = new Label {
        Text = "Missing metadata detected. Optionally fill in the highlighted cells - they are saved on the existing measures / DAX UDFs and used for the new measures.",
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
    };

    var metadataGrid = new DataGridView {
        Dock = DockStyle.Fill,
        AutoGenerateColumns = false,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.CellSelect,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
        BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false,   // required for custom header colours
        EditMode = DataGridViewEditMode.EditOnEnter
    };
    metadataGrid.ColumnHeadersDefaultCellStyle.Font = boldFont;
    metadataGrid.DefaultCellStyle.ForeColor = Color.FromArgb(90, 90, 90);

    Func<string, string, int, bool, DataGridViewTextBoxColumn> addColumn = (name, header, width, editable) =>
    {
        var column = new DataGridViewTextBoxColumn { Name = name, HeaderText = header, Width = width, ReadOnly = !editable };
        if (editable)
        {
            column.DefaultCellStyle = new DataGridViewCellStyle {
                BackColor = editBack, ForeColor = Color.Black, Font = boldFont,
                SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black,
                WrapMode = DataGridViewTriState.True
            };
            column.HeaderCell.Style.BackColor = editHeaderBack;
            column.HeaderCell.Style.ForeColor = Color.Black;
        }
        metadataGrid.Columns.Add(column);
        return column;
    };
    addColumn("Kind", "Type", 80, false);
    addColumn("Name", "Name", 220, false);
    addColumn("Missing", "Missing", 190, false);
    addColumn("FormatString", "Format string (edit)", 190, true);
    addColumn("FormatStringExpression", "Format string expression (edit)", 240, true);
    var descriptionColumn = addColumn("Description", "Description (edit)", 300, true);
    descriptionColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
    descriptionColumn.MinimumWidth = 200;

    // One row per object with something missing; non-applicable cells are greyed and read-only.
    Action<DataGridViewCell> disableCell = cell =>
    {
        cell.ReadOnly = true;
        cell.Style.BackColor = naBack;
        cell.Style.SelectionBackColor = naBack;
        cell.Value = "";
    };
    foreach (var m in Selected.Measures)
    {
        bool noFormat = missingFormatString.Contains(m.Name);
        bool noDescription = missingMeasureDescription.Contains(m.Name);
        if (!noFormat && !noDescription) continue;

        var missing = string.Join(" + ", new[] { noFormat ? "Format string" : null, noDescription ? "Description" : null }.Where(x => x != null));
        int rowIndex = metadataGrid.Rows.Add("Measure", m.Table.Name + "[" + m.Name + "]", missing, "", "", "");
        var row = metadataGrid.Rows[rowIndex];
        row.Tag = m;
        if (!noFormat) { disableCell(row.Cells["FormatString"]); disableCell(row.Cells["FormatStringExpression"]); }
        if (!noDescription) disableCell(row.Cells["Description"]);
    }
    foreach (var f in selectedFunctions.Where(fn => missingFunctionDescription.Contains(fn.Name)))
    {
        int rowIndex = metadataGrid.Rows.Add("DAX UDF", f.Name, "Description", "", "", "");
        var row = metadataGrid.Rows[rowIndex];
        row.Tag = f;
        disableCell(row.Cells["FormatString"]);
        disableCell(row.Cells["FormatStringExpression"]);
    }

    // Suggest common format strings while typing in the Format string column
    metadataGrid.EditingControlShowing += (s, e) =>
    {
        var textBox = e.Control as TextBox;
        if (textBox == null) return;
        if (metadataGrid.CurrentCell.OwningColumn.Name == "FormatString")
        {
            var suggestions = new AutoCompleteStringCollection();
            suggestions.AddRange(new[] {
                "#,0", "#,0.0", "#,0.00", "0", "0.0", "0.00",
                "0%", "0.0%", "#,0.0%;-#,0.0%;#,0.0%", "#,0.00%",
                "#,0;-#,0;#,0", "#,0.0;-#,0.0;#,0.0", "\"kr\" #,0", "#,0 \"kr\""
            });
            suggestions.AddRange(Model.AllMeasures.Select(x => x.FormatString).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToArray());
            textBox.AutoCompleteCustomSource = suggestions;
            textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox.Multiline = false;
        }
        else
        {
            textBox.AutoCompleteMode = AutoCompleteMode.None;
        }
    };

    var metadataHint = new Label {
        Text = "Only the highlighted cells are editable; grey cells are not applicable. Enter EITHER a format string OR a format string expression per measure, not both. " +
               "Blank cells are left unchanged. Suggestions are shown while typing a format string.",
        AutoSize = true, MaximumSize = new Size(1100, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
    };

    var metadataButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var applyAndContinue = new Button { Text = "Apply changes and continue", AutoSize = true, Font = boldFont, Margin = new Padding(0, 0, 8, 0) };
    var continueUnchanged = new Button { Text = "Continue without changes", AutoSize = true, Margin = new Padding(0, 0, 8, 0), DialogResult = DialogResult.Ignore };
    var metadataCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    metadataButtons.Controls.AddRange(new Control[] { applyAndContinue, continueUnchanged, metadataCancel });

    applyAndContinue.Click += (s, e) =>
    {
        metadataGrid.EndEdit();
        foreach (DataGridViewRow row in metadataGrid.Rows)
        {
            var fs  = (Convert.ToString(row.Cells["FormatString"].Value) ?? "").Trim();
            var fse = (Convert.ToString(row.Cells["FormatStringExpression"].Value) ?? "").Trim();
            if (fs.Length > 0 && fse.Length > 0)
            {
                MessageBox.Show("Enter either a format string or a format string expression for " + row.Cells["Name"].Value + ", not both.",
                    "Format string", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                metadataGrid.CurrentCell = row.Cells["FormatStringExpression"];
                return;
            }
        }
        metadataForm.DialogResult = DialogResult.OK;
        metadataForm.Close();
    };

    metadataLayout.Controls.Add(metadataHeading, 0, 0);
    metadataLayout.Controls.Add(metadataGrid, 0, 1);
    metadataLayout.Controls.Add(metadataHint, 0, 2);
    metadataLayout.Controls.Add(metadataButtons, 0, 3);
    // No AcceptButton: Enter commits the cell being edited instead of closing.
    metadataForm.CancelButton = metadataCancel;
    metadataForm.Shown += (s, e) =>
    {
        foreach (DataGridViewRow row in metadataGrid.Rows)
            foreach (DataGridViewCell cell in row.Cells)
                if (!cell.ReadOnly) { metadataGrid.CurrentCell = cell; metadataGrid.BeginEdit(true); return; }
    };

    var metadataResult = metadataForm.ShowDialog();
    if (metadataResult == DialogResult.Cancel)
        return;

    if (metadataResult == DialogResult.OK)
    {
        int updated = 0;
        foreach (DataGridViewRow row in metadataGrid.Rows)
        {
            var fs   = (Convert.ToString(row.Cells["FormatString"].Value) ?? "").Trim();
            var fse  = (Convert.ToString(row.Cells["FormatStringExpression"].Value) ?? "").Trim();
            var desc = (Convert.ToString(row.Cells["Description"].Value) ?? "").Trim();

            var measure = row.Tag as Measure;
            if (measure != null)
            {
                if (fs.Length > 0)  { measure.FormatString = fs; updated++; }
                if (fse.Length > 0) { measure.FormatStringExpression = fse; updated++; }
                if (desc.Length > 0) { measure.Description = desc; updated++; }
            }
            var function = row.Tag as Function;
            if (function != null && desc.Length > 0) { function.Description = desc; updated++; }
        }
        if (updated > 0)
            Info(updated + " metadata value(s) saved on the existing measures / DAX UDFs. The new measures will use them.");
    }
}

// Helper: prompts for the measure naming (Prefix and/or Suffix - at least one required, since
// Cancel/blank on both is not a valid name), and/or an optional Folder name to append - showing a
// live example based on a sample measure (the first selected measure) as the user types. Only the
// section(s) actually needed (needMeasureNaming / needFolderNaming) are shown.
Func<Function, bool, bool, Measure, (string prefix, string suffix, string folderName)> ShowFunctionNamingDialog = (f, needMeasureNaming, needFolderNaming, sampleMeasure) =>
{
    var form = new Form
    {
        Text = $"Define naming for function '{f.Name}'",
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
        measurePreview.Text = $"Example (for '{sampleMeasureName}'): {example}";
    };
    Action updateFolderPreview = () =>
    {
        if (folderPreview == null) return;
        string folderTxt = folderBox.Text.Trim();
        string example = string.IsNullOrWhiteSpace(folderTxt)
            ? sampleFolder + "\\"
            : (sampleFolder + "\\" + folderTxt).Trim('\\');
        folderPreview.Text = $"Example folder (for '{sampleMeasureName}'): {example}";
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
        return ("", "", "");

    return (
        prefixBox != null ? prefixBox.Text.Trim() : "",
        suffixBox != null ? suffixBox.Text.Trim() : "",
        folderBox != null ? folderBox.Text.Trim() : ""
    );
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
            functionMeasurePrefixes[f.Name] = naming.prefix;
            functionMeasureSuffixes[f.Name] = naming.suffix;
        }
        if (needsFolderNaming)
        {
            functionFolderNames[f.Name] = naming.folderName;
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
            alreadyExistingMeasures.Add($"{newMeasureName} (in {m.Table.Name})");
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