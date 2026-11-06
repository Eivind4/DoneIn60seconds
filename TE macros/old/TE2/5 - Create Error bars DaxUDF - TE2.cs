// ADD ERROR BAR HELPER MEASURES AS DAX UDFs
// ============================================
//
// Purpose: Generate two reusable, model-DEPENDENT DAX User Defined Functions (UDFs) used to drive
// the "Upper" error bar value of a visual comparing an Actual measure (measure1) against a Target
// measure (measure2, e.g. Budget/Plan/Prior period). Each function returns the Actual measure's
// value only when the % difference between Actual and Target crosses a threshold - otherwise it
// returns BLANK() (so the error bar simply doesn't show).
//
//   - Local.ErrorBar.ActualAboveTarget
//       Returns measure1 (Actual) when it is more than the UPPER threshold above measure2 (Target).
//       Add this measure as the "Upper" error bar value on the TARGET (measure2) series.
//   - Local.ErrorBar.ActualBelowTarget
//       Returns measure1 (Actual) when it is more than the LOWER threshold below measure2 (Target).
//       Add this measure as the "Upper" error bar value on the ACTUAL (measure1) series.
//
// Both functions take exactly two ANYREF parameters (actual, target) - the same shape as the
// Comparison.* functions - so they can be turned into concrete measures using
// "5 - Compare first selected measure with second selected measure.cs" (that macro has been
// updated to also recognize functions with "ErrorBar" in the name).
//
// The upper/lower threshold used inside each function can be sourced the same way as
// "5 - Create Conditional formatting DaxUDF.cs":
//   - Fixed: hard-coded literal values you type in below.
//   - Dynamic: either (a) pick existing Upper/Lower bound measures already in the model, or
//              (b) create a new "Conditional Format Input" calculated table + Max/Min measures
//                  (identical logic/table to the Conditional Formatting macro - reused as-is if
//                  it already exists).
//
// Per the user's request, the "# Rows" / budget-branching logic seen in some hand-written error
// bar measures is intentionally NOT included here - add that yourself afterwards if needed.
//
// Safety checks performed before creating each function:
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function is reused/referenced instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one, Overwrite the
//       existing (outdated) definition, or Skip creating it (never silently overwritten).
//
#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;
using System.Globalization;

// TE2 COMPATIBILITY NOTE - changes made when porting this script from TE3 to TE2:
// 1. Converted all 21 $"..." / $@"..." interpolated strings to string concatenation (TE2's
//    compiler doesn't support string interpolation). The 2 multi-line $@"..." verbatim DAX-body
//    strings were converted to concatenated @"..." verbatim segments (keeping real line breaks),
//    not regular strings, to avoid illegal embedded raw newlines.
// 2. Replaced SelectUpperLowerMeasures's ValueTuple return type "(Measure upper, Measure lower)?"
//    with a plain "Measure[]" 2-element array (index 0 = upper, index 1 = lower) - TE2's older
//    compiler does not support C# 7 ValueTuples.
// No other changes were needed (this script already used Func<> delegates rather than local
// functions, and had no other TE2-incompatible constructs).

Application.UseWaitCursor = false;
WaitFormVisible = false;

Func<string, string> NormalizeExpr = s => string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
Func<decimal, string> FormatDecimalLiteral = d => d.ToString("0.####", CultureInfo.InvariantCulture);

// ================================================================================================
// DIALOG 1: Fixed vs Dynamic threshold source, plus the upper/lower values (or, for Dynamic, the
// choice between picking existing measures or creating a new calculated table), and whether the
// values represent percentage points (e.g. entering 15 for 15%).
// ================================================================================================
var inputForm = new Form
{
    Text = "Error bars - threshold input source",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 520,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var inputFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
inputForm.Controls.Add(inputFlow);
Action<string> inputLabel = s => inputFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

inputLabel("Choose how the upper/lower error bar threshold is determined:");
var fixedRadio   = new RadioButton { Text = "Fixed threshold input (hard-coded values)", AutoSize = true, Checked = true };
var dynamicRadio = new RadioButton { Text = "Dynamic (existing measures, or a new calculated table)", AutoSize = true };
inputFlow.Controls.Add(fixedRadio);
inputFlow.Controls.Add(dynamicRadio);

// -- Fixed panel --
var fixedPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0) };
fixedPanel.Controls.Add(new Label { Text = "Lower threshold (ActualBelowTarget fires below this):", AutoSize = true });
var minFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = -15 };
fixedPanel.Controls.Add(minFixedInput);
fixedPanel.Controls.Add(new Label { Text = "Upper threshold (ActualAboveTarget fires above this):", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = 15 };
fixedPanel.Controls.Add(maxFixedInput);
inputFlow.Controls.Add(fixedPanel);

// -- Dynamic panel --
var dynamicPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0), Visible = false };
var pickRadio   = new RadioButton { Text = "Pick existing Upper/Lower bound measures already in the model", AutoSize = true, Checked = true };
var createRadio = new RadioButton { Text = "Create a new 'Conditional Format Input' calculated table", AutoSize = true };
dynamicPanel.Controls.Add(pickRadio);
dynamicPanel.Controls.Add(createRadio);

var createTablePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0), Visible = false };
createTablePanel.Controls.Add(new Label { Text = "Table minimum:", AutoSize = true });
var minDynInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, Value = -30 };
createTablePanel.Controls.Add(minDynInput);
createTablePanel.Controls.Add(new Label { Text = "Table maximum:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxDynInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, Value = 30 };
createTablePanel.Controls.Add(maxDynInput);
createTablePanel.Controls.Add(new Label { Text = "Increment:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var incDynInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 10000, Value = 1 };
createTablePanel.Controls.Add(incDynInput);
dynamicPanel.Controls.Add(createTablePanel);
inputFlow.Controls.Add(dynamicPanel);

var percentCheckbox = new CheckBox { Text = "Values represent percentage points (e.g. enter 15 for 15% - divided by 100)", AutoSize = true, Checked = true, Margin = new Padding(0, 10, 0, 0) };
inputFlow.Controls.Add(percentCheckbox);

Action updateInputVis = () =>
{
    fixedPanel.Visible = fixedRadio.Checked;
    dynamicPanel.Visible = dynamicRadio.Checked;
    createTablePanel.Visible = dynamicRadio.Checked && createRadio.Checked;
};
fixedRadio.CheckedChanged   += (s, e) => updateInputVis();
dynamicRadio.CheckedChanged += (s, e) => updateInputVis();
pickRadio.CheckedChanged    += (s, e) => updateInputVis();
createRadio.CheckedChanged  += (s, e) => updateInputVis();

var inputBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var inputOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var inputCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
inputBtnPanel.Controls.Add(inputOk);
inputBtnPanel.Controls.Add(inputCancel);
inputFlow.Controls.Add(inputBtnPanel);
inputForm.AcceptButton = inputOk;
inputForm.CancelButton = inputCancel;

if (inputForm.ShowDialog() == DialogResult.Cancel) return;

bool isDynamic      = dynamicRadio.Checked;
bool isPercentage   = percentCheckbox.Checked;
bool pickExisting   = isDynamic && pickRadio.Checked;
bool createNewTable = isDynamic && createRadio.Checked;

// ================================================================================================
// Measure-picker helpers (only needed for the "pick existing measures" Dynamic option) - a search +
// tree-view dual picker, matching the one used in
// "5 - Compare first selected measure with second selected measure.cs".
// ================================================================================================
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

// Wires up a TreeView's checkboxes so that checking a measure leaf node acts as a single-select
// (any other checked leaf in the same tree is automatically unchecked), while checking a
// non-leaf (table/folder) node is not a valid selection and is reverted.
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
            e.Node.Checked = false;
            suppress = false;
            return;
        }

        if (e.Node.Checked)
        {
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

// Shows a single dialog with two independent search + tree-view measure pickers side by side
// ("Upper bound" on the left, "Lower bound" on the right) and returns the chosen pair as a
// 2-element array (index 0 = upper, index 1 = lower), or null if cancelled.
// TE2 note: uses a Measure[] instead of a ValueTuple (TE2's older compiler does not support tuples).
Func<Measure[]> SelectUpperLowerMeasures = () =>
{
    var allMeasures = Model.AllMeasures.ToList();

    var pickForm = new Form { Text = "Select Upper and Lower bound measures", Width = 1000, Height = 640, StartPosition = FormStartPosition.CenterScreen };

    Measure selectedUpper = null;
    Measure selectedLower = null;

    Func<string, int, int, int, Label> BuildLabel = (text, x, y, w) =>
        new Label { Text = text, Location = new Point(x, y), Width = w, Height = 20, Font = new Font(pickForm.Font, FontStyle.Bold) };

    // ── Upper (left) ──
    var upperLabel    = BuildLabel("Upper bound measure (ActualAboveTarget threshold):", 20, 15, 440);
    var upperSearch   = new TextBox { Location = new Point(20, 40), Width = 440 };
    var upperTree     = new TreeView { Location = new Point(20, 68), Width = 440, Height = 450, CheckBoxes = true };
    var upperSelLabel = new Label { Text = "Selected: (none)", Location = new Point(20, 522), Width = 440, Height = 20, Font = new Font(pickForm.Font, FontStyle.Italic) };

    // ── Lower (right) ──
    var lowerLabel    = BuildLabel("Lower bound measure (ActualBelowTarget threshold):", 500, 15, 440);
    var lowerSearch   = new TextBox { Location = new Point(500, 40), Width = 440 };
    var lowerTree     = new TreeView { Location = new Point(500, 68), Width = 440, Height = 450, CheckBoxes = true };
    var lowerSelLabel = new Label { Text = "Selected: (none)", Location = new Point(500, 522), Width = 440, Height = 20, Font = new Font(pickForm.Font, FontStyle.Italic) };

    WireSingleCheckSelection(upperTree, upperSelLabel, m => selectedUpper = m);
    WireSingleCheckSelection(lowerTree, lowerSelLabel, m => selectedLower = m);

    RebuildMeasureTree(upperTree, "", allMeasures, () => selectedUpper);
    RebuildMeasureTree(lowerTree, "", allMeasures, () => selectedLower);

    upperSearch.TextChanged += (s, e) => RebuildMeasureTree(upperTree, upperSearch.Text, allMeasures, () => selectedUpper);
    lowerSearch.TextChanged += (s, e) => RebuildMeasureTree(lowerTree, lowerSearch.Text, allMeasures, () => selectedLower);

    var okButton     = new Button { Text = "OK",     Location = new Point(380, 560), Width = 80, Height = 30 };
    var cancelButton = new Button { Text = "Cancel", Location = new Point(480, 560), Width = 80, Height = 30, DialogResult = DialogResult.Cancel };

    okButton.Click += (s, e) =>
    {
        if (selectedUpper == null || selectedLower == null)
        {
            MessageBox.Show("Please select both an Upper and a Lower bound measure.", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        pickForm.DialogResult = DialogResult.OK;
        pickForm.Close();
    };

    pickForm.Controls.Add(upperLabel);
    pickForm.Controls.Add(upperSearch);
    pickForm.Controls.Add(upperTree);
    pickForm.Controls.Add(upperSelLabel);
    pickForm.Controls.Add(lowerLabel);
    pickForm.Controls.Add(lowerSearch);
    pickForm.Controls.Add(lowerTree);
    pickForm.Controls.Add(lowerSelLabel);
    pickForm.Controls.Add(okButton);
    pickForm.Controls.Add(cancelButton);

    pickForm.AcceptButton = okButton;
    pickForm.CancelButton = cancelButton;

    if (pickForm.ShowDialog() != DialogResult.OK)
        return null;

    return new Measure[] { selectedUpper, selectedLower };
};

// ================================================================================================
// Resolve the upper/lower threshold DAX references used inside the two functions - either literal
// fixed values, references to existing measures, or references to a new/reused calculated table.
// ================================================================================================
string upperArg, lowerArg;
var notes = new List<string>();

if (!isDynamic)
{
    decimal minRaw = minFixedInput.Value;
    decimal maxRaw = maxFixedInput.Value;
    decimal minDecimal = isPercentage ? minRaw / 100 : minRaw;
    decimal maxDecimal = isPercentage ? maxRaw / 100 : maxRaw;

    lowerArg = FormatDecimalLiteral(minDecimal);
    upperArg = FormatDecimalLiteral(maxDecimal);
}
else if (pickExisting)
{
    var picked = SelectUpperLowerMeasures();
    if (picked == null)
    {
        Info("Cancelled - no bound measures were selected.");
        return;
    }
    upperArg = picked[0].DaxObjectFullName;
    lowerArg = picked[1].DaxObjectFullName;
}
else // createNewTable
{
    const string inputTableName = "Conditional Format Input";
    var existingTable = Model.Tables.FirstOrDefault(t => t.Name == inputTableName);

    string maxMeasureName = isPercentage ? "Max Value percent" : "Max Value";
    string minMeasureName = isPercentage ? "Min Value percent" : "Min Value";

    bool createTableAndMeasures = true;

    if (existingTable != null)
    {
        var choice = MessageBox.Show(
            "A table named '" + (inputTableName) + "' already exists.\n\n" +
            "Yes = Overwrite (delete and recreate it with the values above)\n" +
            "No = Reuse the existing table as-is (values above are ignored)\n" +
            "Cancel = Abort this script",
            "Conditional Format Input table already exists",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (choice == DialogResult.Cancel) return;

        if (choice == DialogResult.Yes)
        {
            existingTable.Delete();
            existingTable = null;
        }
        else
        {
            createTableAndMeasures = false;
        }
    }

    if (createTableAndMeasures)
    {
        int min = (int)minDynInput.Value;
        int max = (int)maxDynInput.Value;
        int inc = (int)incDynInput.Value;

        var table = Model.AddCalculatedTable(inputTableName, "GENERATESERIES(" + (min) + ", " + (max) + ", " + (inc) + ")");
        table.Description = "To set the dynamic max and min value for conditional formatting / error bars";

        // TE3 auto-creates the "Value" column from GENERATESERIES; fall back to the first column
        // in case a differently-named/typed column is inferred.
        var valueColumn = table.Columns.Contains("Value") ? table.Columns["Value"] : table.Columns.First();
        string valueColRef = valueColumn.DaxObjectFullName;
        string tableRef    = table.DaxObjectFullName;

        var valueMeasure = table.AddMeasure(
            (inputTableName) + " Value",
            "SELECTEDVALUE(" + (valueColRef) + ", 0)",
            "Conditional Formatting");
        valueMeasure.Description = "Selected value from the " + (inputTableName) + " table, defaults to 0 if nothing selected";
        valueMeasure.FormatString = "#,0";

        string divideSuffix = isPercentage ? " / 100" : "";
        string formatString = isPercentage ? "#,0.0%;-#,0.0%;#,0.0%" : "#,0";

        var maxMeasure = table.AddMeasure(
            maxMeasureName,
            "MAXX(" + (tableRef) + ", " + (valueColRef) + ")" + (divideSuffix),
            "Conditional Formatting");
        maxMeasure.Description = isPercentage
            ? "Maximum value from " + (inputTableName) + " table divided by 100 to convert to decimal"
            : "Maximum value from " + (inputTableName) + " table";
        maxMeasure.FormatString = formatString;

        var minMeasure = table.AddMeasure(
            minMeasureName,
            "MINX(" + (tableRef) + ", " + (valueColRef) + ")" + (divideSuffix),
            "Conditional Formatting");
        minMeasure.Description = isPercentage
            ? "Minimum value from " + (inputTableName) + " table divided by 100 to convert to decimal"
            : "Minimum value from " + (inputTableName) + " table";
        minMeasure.FormatString = formatString;

        existingTable = table;
    }

    var maxMeasureRef = existingTable.Measures.FirstOrDefault(m => m.Name == maxMeasureName);
    var minMeasureRef = existingTable.Measures.FirstOrDefault(m => m.Name == minMeasureName);
    if (maxMeasureRef == null || minMeasureRef == null)
    {
        Error("Could not find measures '" + (maxMeasureName) + "' / '" + (minMeasureName) + "' on the existing '" + (inputTableName) + "' table. Re-run and choose Overwrite, or align the percentage option with how the table was originally created.");
        return;
    }

    upperArg = maxMeasureRef.DaxObjectFullName;
    lowerArg = minMeasureRef.DaxObjectFullName;
}

// ================================================================================================
// DIALOG 2: Review/edit the suggested function names.
// ================================================================================================
var nameForm = new Form
{
    Text = "Error bars - function names",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 520,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var nameFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
nameForm.Controls.Add(nameFlow);
nameFlow.Controls.Add(new Label { Text = "Review/edit the function names to be created:", AutoSize = true, Margin = new Padding(0, 0, 0, 10) });

var aboveRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
aboveRow.Controls.Add(new Label { Text = "Above Target:", Width = 110, Margin = new Padding(0, 4, 0, 0) });
var aboveNameBox = new TextBox { Width = 320, Text = "Local.ErrorBar.ActualAboveTarget" };
aboveRow.Controls.Add(aboveNameBox);
nameFlow.Controls.Add(aboveRow);

var belowRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
belowRow.Controls.Add(new Label { Text = "Below Target:", Width = 110, Margin = new Padding(0, 4, 0, 0) });
var belowNameBox = new TextBox { Width = 320, Text = "Local.ErrorBar.ActualBelowTarget" };
belowRow.Controls.Add(belowNameBox);
nameFlow.Controls.Add(belowRow);

var nameBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
var nameOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var nameCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
nameBtnPanel.Controls.Add(nameOk);
nameBtnPanel.Controls.Add(nameCancel);
nameFlow.Controls.Add(nameBtnPanel);
nameForm.AcceptButton = nameOk;
nameForm.CancelButton = nameCancel;

if (nameForm.ShowDialog() == DialogResult.Cancel) return;

string aboveName = string.IsNullOrWhiteSpace(aboveNameBox.Text) ? "Local.ErrorBar.ActualAboveTarget" : aboveNameBox.Text.Trim();
string belowName = string.IsNullOrWhiteSpace(belowNameBox.Text) ? "Local.ErrorBar.ActualBelowTarget" : belowNameBox.Text.Trim();

// ================================================================================================
// Create the two functions - same skip/rename/overwrite conflict handling used by the other DaxUDF
// generator macros. No annotations are added, since these functions are meant to be consumed
// directly by "5 - Compare first selected measure with second selected measure.cs" (which only
// needs the two ANYREF parameters), not by "5 - Create measure from model dependent functions.cs".
// ================================================================================================
Action<string, string, string> CreateUDF = (funcName, daxBody, description) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        notes.Add("'" + (funcName) + "' was NOT created - using the existing function '" + (dupExprFunc.Name) + "' instead, since it has an identical expression.");
        return;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            notes.Add("'" + (funcName) + "' already exists with the same expression - left unchanged.");
            return;
        }

        var conflictDlg = new Form { Text = "Function name already exists", Width = 480, Height = 250, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false };
        var conflictFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = "A function named '" + (funcName) + "' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(440, 0), Margin = new Padding(0, 0, 0, 10) });

        var renameRow   = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var renameRadio = new RadioButton { Text = "Rename to:", AutoSize = true, Checked = true, Margin = new Padding(0, 3, 8, 0) };
        var renameBox   = new TextBox { Width = 250, Text = funcName };
        renameRow.Controls.Add(renameRadio);
        renameRow.Controls.Add(renameBox);
        conflictFlow.Controls.Add(renameRow);

        var overwriteRadio = new RadioButton { Text = "Overwrite the existing definition with this new one", AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var skipRadio       = new RadioButton { Text = "Skip - do not create this UDF", AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        conflictFlow.Controls.Add(overwriteRadio);
        conflictFlow.Controls.Add(skipRadio);

        var conflictBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
        var conflictOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
        conflictBtnPanel.Controls.Add(conflictOk);
        conflictFlow.Controls.Add(conflictBtnPanel);
        conflictDlg.AcceptButton = conflictOk;

        conflictDlg.ShowDialog();

        if (overwriteRadio.Checked)
        {
            existing.Expression  = daxBody;
            existing.Description = description;
            notes.Add("'" + (funcName) + "' already existed with a different expression - its definition was OVERWRITTEN with the new one.");
            return;
        }
        else if (skipRadio.Checked)
        {
            notes.Add("'" + (funcName) + "' was NOT created - a function with that name already exists (existing definition kept).");
            return;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                notes.Add("'" + (funcName) + "' was NOT created - no valid new name was provided (existing definition kept).");
                return;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                notes.Add("'" + (funcName) + "' was NOT created - the alternate name '" + (newName) + "' also already exists.");
                return;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
};

// "measure1"/"measure2" are left as literal placeholder text in the descriptions below on purpose
// - "5 - Compare first selected measure with second selected measure.cs" replaces them with the
// actual/target measure's fully-qualified name when it turns this function into a real measure.
string aboveBody = @"(
    // What you are measuring - the Actual value (measure1)
    actual: ANYREF,
    // What you are comparing to - the Target/Budget/Plan value (measure2)
    target: ANYREF
) =>
    VAR _Actual = actual
    VAR _Target = target
    VAR _PctChange = DIVIDE(_Actual - _Target, ABS(_Target))
    RETURN
        IF(
            _PctChange > " + (upperArg) + @",
            _Actual
        )";

string belowBody = @"(
    // What you are measuring - the Actual value (measure1)
    actual: ANYREF,
    // What you are comparing to - the Target/Budget/Plan value (measure2)
    target: ANYREF
) =>
    VAR _Actual = actual
    VAR _Target = target
    VAR _PctChange = DIVIDE(_Actual - _Target, ABS(_Target))
    RETURN
        IF(
            _PctChange < " + (lowerArg) + @",
            _Actual
        )";

string aboveDescription = "Error bar helper: returns measure1 only when it is more than the upper threshold above measure2 (otherwise blank). Add the resulting measure as the 'Upper' error bar value on the TARGET (measure2) series.";
string belowDescription = "Error bar helper: returns measure1 only when it is more than the lower threshold below measure2 (otherwise blank). Add the resulting measure as the 'Upper' error bar value on the ACTUAL (measure1) series.";

CreateUDF(aboveName, aboveBody, aboveDescription);
CreateUDF(belowName, belowBody, belowDescription);

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = "Done! Created/updated the ErrorBar functions.\n\n" +
    "Next step: run '5 - Compare first selected measure with second selected measure.cs', pick your Actual and Target measures, and select the ErrorBar function(s) above to generate the concrete error bar measures.";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n => "- " + n));

Info(doneSummary);
