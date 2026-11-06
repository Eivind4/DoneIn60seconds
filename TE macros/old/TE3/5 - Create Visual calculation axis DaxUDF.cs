// ADD VISUAL CALCULATION AXIS HELPERS AS DAX UDFs
// =================================================
//
// Purpose: Generate reusable DAX User Defined Functions (UDFs) for setting a dynamic axis
// min/max in Power BI "visual calculations" (works for either the X- or Y-axis), for 1, 2 or 3
// input measures/columns. For each selected input count, BOTH a "Lower" and an "Upper" bound
// function is generated - mirroring the existing Lower3_columns / Upper3_columns pattern, which
// was previously missing for One/Two inputs:
//   - Local.VisualCalc.Axis.OneInputLower    / OneInputUpper
//   - Local.VisualCalc.Axis.TwoInputsLower   / TwoInputsUpper
//   - Local.VisualCalc.Axis.ThreeInputsLower / ThreeInputsUpper
//
// Each function applies a padding margin (default 20% extra) on top of the min/max of the
// selected inputs across the visual matrix (via ROWS / EXPANDALL), and optionally clamps
// ("min/max restricted") the result to a fixed bound (default -1 / 1) so the axis doesn't blow out
// when values can exceed +/-100% - matching the WINDOW/ABS clamp used in Lower3_columns/Upper3_columns.
// If not restricted, the simpler pattern from DynYaxis.OneInput/TwoInputs is used instead (only
// ensures 0 is included in the axis range).
//
// Safety checks performed before creating each UDF (same pattern as the other DaxUDF macros):
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function is reused/kept as-is instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one, Overwrite the
//       existing (outdated) definition, or Skip creating it (never silently overwritten).
//
#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;
using System.Globalization;

Application.UseWaitCursor = false;
WaitFormVisible = false;

Func<string, string> NormalizeExpr = s => string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
Func<decimal, string> FormatDecimalLiteral = d => d.ToString("0.####", CultureInfo.InvariantCulture);

// ================================================================================================
// DIALOG 1: How many inputs (1/2/3 - multiple can be selected), padding margin, optional min/max
// restriction (and its bounds), and the function name prefix.
// ================================================================================================
var mainForm = new Form
{
    Text = "Visual calculation axis - options",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 500,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var mainFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
mainForm.Controls.Add(mainFlow);
Action<string> mainLabel = s => mainFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

mainLabel("Number of input measures/columns to support (select one or more):");
var oneInputCheck   = new CheckBox { Text = "1 input",  AutoSize = true, Checked = true };
var twoInputsCheck  = new CheckBox { Text = "2 inputs", AutoSize = true, Checked = true };
var threeInputsCheck = new CheckBox { Text = "3 inputs", AutoSize = true, Checked = true };
mainFlow.Controls.Add(oneInputCheck);
mainFlow.Controls.Add(twoInputsCheck);
mainFlow.Controls.Add(threeInputsCheck);

mainLabel("Padding margin (extra % added on top of the min/max value, e.g. 20 = 1.2x):");
var paddingInput = new NumericUpDown { Width = 100, Minimum = 0, Maximum = 500, DecimalPlaces = 1, Increment = 5, Value = 20 };
mainFlow.Controls.Add(paddingInput);

var restrictCheck = new CheckBox { Text = "Restrict (clamp) the axis to a min/max bound - recommended so the axis doesn't blow out when values can exceed +/-100%", AutoSize = true, MaximumSize = new Size(460, 0), Checked = true, Margin = new Padding(0, 10, 0, 0) };
mainFlow.Controls.Add(restrictCheck);

var restrictPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0) };
restrictPanel.Controls.Add(new Label { Text = "Minimum bound (used for the Lower functions):", AutoSize = true });
var minBoundInput = new NumericUpDown { Width = 100, Minimum = -1000, Maximum = 1000, DecimalPlaces = 2, Increment = 0.1M, Value = -1 };
restrictPanel.Controls.Add(minBoundInput);
restrictPanel.Controls.Add(new Label { Text = "Maximum bound (used for the Upper functions):", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxBoundInput = new NumericUpDown { Width = 100, Minimum = -1000, Maximum = 1000, DecimalPlaces = 2, Increment = 0.1M, Value = 1 };
restrictPanel.Controls.Add(maxBoundInput);
mainFlow.Controls.Add(restrictPanel);

restrictCheck.CheckedChanged += (s, e) => restrictPanel.Visible = restrictCheck.Checked;

mainLabel("(Optional) Function name prefix  [default: Local.VisualCalc.Axis]:");
var prefixTextbox = new TextBox { Width = 350, Text = "Local.VisualCalc.Axis" };
mainFlow.Controls.Add(prefixTextbox);

var mainBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var mainOk     = new Button { Text = "OK",     AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
var mainCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
mainBtnPanel.Controls.Add(mainOk);
mainBtnPanel.Controls.Add(mainCancel);
mainFlow.Controls.Add(mainBtnPanel);
mainForm.AcceptButton = mainOk;
mainForm.CancelButton = mainCancel;

// No DialogResult set on mainOk directly (validated manually here instead), so an empty
// input-count selection doesn't close the dialog.
mainOk.Click += (s, e) =>
{
    if (!oneInputCheck.Checked && !twoInputsCheck.Checked && !threeInputsCheck.Checked)
    {
        MessageBox.Show("Please select at least one input count (1, 2 or 3).", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }
    mainForm.DialogResult = DialogResult.OK;
    mainForm.Close();
};

if (mainForm.ShowDialog() == DialogResult.Cancel) return;

var selectedCounts = new List<int>();
if (oneInputCheck.Checked)    selectedCounts.Add(1);
if (twoInputsCheck.Checked)   selectedCounts.Add(2);
if (threeInputsCheck.Checked) selectedCounts.Add(3);

decimal paddingPercent   = paddingInput.Value;
decimal paddingMultiplier = 1 + (paddingPercent / 100);
string paddingLiteral    = FormatDecimalLiteral(paddingMultiplier);

bool isRestricted = restrictCheck.Checked;
decimal minBound = minBoundInput.Value;
decimal maxBound = maxBoundInput.Value;
string minBoundLiteral = FormatDecimalLiteral(minBound);
string maxBoundLiteral = FormatDecimalLiteral(maxBound);

string namePrefix = string.IsNullOrWhiteSpace(prefixTextbox.Text) ? "Local.VisualCalc.Axis" : prefixTextbox.Text.Trim();

// ================================================================================================
// Helper builders: parameter comment/name for each input count, the combined MIN/MAX expression
// over the selected inputs, a suggested function name, the DAX body, and the description.
// ================================================================================================
string[] inputDescriptions = {
    "First value (e.g. actual measure or column)",
    "Second value (e.g. target/forecast measure or column)",
    "Third value (e.g. additional measure or column)"
};

Func<int, string> ParamsBlock = n =>
{
    var lines = new List<string>();
    for (int i = 1; i <= n; i++)
    {
        lines.Add($"        // {inputDescriptions[i - 1]}");
        lines.Add($"        input{i}: ANYREF" + (i < n ? "," : ""));
    }
    return string.Join("\r\n", lines);
};

Func<int, string> CombinedMin = n => n == 1 ? "input1" : n == 2 ? "MIN(input1, input2)" : "MIN(MIN(input1, input2), input3)";
Func<int, string> CombinedMax = n => n == 1 ? "input1" : n == 2 ? "MAX(input1, input2)" : "MAX(MAX(input1, input2), input3)";

Func<int, string> InputCountWord = n => n == 1 ? "OneInput" : n == 2 ? "TwoInputs" : "ThreeInputs";

Func<int, bool, string> SuggestName = (n, isUpper) =>
    $"{namePrefix}.{InputCountWord(n)}{(isUpper ? "Upper" : "Lower")}";

Func<int, bool, string> BuildAxisBody = (n, isUpper) =>
{
    string paramsBlock = ParamsBlock(n);
    string combined    = isUpper ? CombinedMax(n) : CombinedMin(n);

    if (!isRestricted)
    {
        string aggFunc  = isUpper ? "MAXX" : "MINX";
        string clampFn  = isUpper ? "MAX"  : "MIN";
        string varName  = isUpper ? "_Max" : "_Min";
        return
$@"(
{paramsBlock}
    ) =>
    VAR {varName} = CALCULATE({aggFunc}(ROWS, {combined}), EXPANDALL(ROWS))
    VAR _Adjusted = {varName} * {paddingLiteral}
    RETURN {clampFn}(0, _Adjusted)";
    }
    else
    {
        string aggFunc    = isUpper ? "MAXX" : "MINX";
        string outerClamp = isUpper ? "MIN"  : "MAX";
        string innerClamp = isUpper ? "MAX"  : "MIN";
        string bound      = isUpper ? maxBoundLiteral : minBoundLiteral;
        return
$@"(
{paramsBlock}
    ) =>
    EXPANDALL(
        {outerClamp}(
            {bound},
            {innerClamp}(
                0,
                {aggFunc}(
                    WINDOW(1, ABS, -1, ABS, ROWS),
                    {combined} * {paddingLiteral}
                )
            )
        ),
        ROWS
    )";
    }
};

Func<int, bool, string> BuildDescription = (n, isUpper) =>
{
    string boundWord = isUpper ? "Upper" : "Lower";
    string plural     = n == 1 ? "input" : "inputs";
    string desc = $"For visual calculations axis ({boundWord} bound, {n} {plural}) - includes a {paddingPercent}% padding margin";
    desc += isRestricted
        ? $", min/max restricted to {(isUpper ? maxBoundLiteral : minBoundLiteral)}."
        : ".";
    return desc;
};

// ================================================================================================
// DIALOG 2: Preview / edit the suggested names for each function about to be created.
// ================================================================================================
var combos = new List<(int n, bool isUpper)>();
foreach (var n in selectedCounts)
{
    combos.Add((n, false)); // Lower
    combos.Add((n, true));  // Upper
}

var nameForm = new Form
{
    Text = "Visual calculation axis - function names",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 560,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var nameFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
nameForm.Controls.Add(nameFlow);
nameFlow.Controls.Add(new Label { Text = "Review/edit the function name for each UDF to be created:", AutoSize = true, Margin = new Padding(0, 0, 0, 10) });

var nameBoxes = new List<TextBox>();
foreach (var (n, isUpper) in combos)
{
    string plural = n == 1 ? "input" : "inputs";
    var row = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
    row.Controls.Add(new Label { Text = $"{n} {plural}, {(isUpper ? "Upper" : "Lower")}:", Width = 140, Margin = new Padding(0, 4, 0, 0) });
    var box = new TextBox { Width = 340, Text = SuggestName(n, isUpper) };
    row.Controls.Add(box);
    nameFlow.Controls.Add(row);
    nameBoxes.Add(box);
}

var nameBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
var nameOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var nameCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
nameBtnPanel.Controls.Add(nameOk);
nameBtnPanel.Controls.Add(nameCancel);
nameFlow.Controls.Add(nameBtnPanel);
nameForm.AcceptButton = nameOk;
nameForm.CancelButton = nameCancel;

if (nameForm.ShowDialog() == DialogResult.Cancel) return;

// ================================================================================================
// Create the functions - same skip/rename/overwrite conflict handling used by the other DaxUDF
// generator macros.
// ================================================================================================
var notes = new List<string>();

Action<string, string, string> CreateUDF = (funcName, daxBody, description) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        notes.Add($"'{funcName}' was NOT created - using the existing function '{dupExprFunc.Name}' instead, since it has an identical expression.");
        return;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            notes.Add($"'{funcName}' already exists with the same expression - left unchanged.");
            return;
        }

        var conflictDlg = new Form { Text = "Function name already exists", Width = 480, Height = 250, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false };
        var conflictFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = $"A function named '{funcName}' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(440, 0), Margin = new Padding(0, 0, 0, 10) });

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
            notes.Add($"'{funcName}' already existed with a different expression - its definition was OVERWRITTEN with the new one.");
            return;
        }
        else if (skipRadio.Checked)
        {
            notes.Add($"'{funcName}' was NOT created - a function with that name already exists (existing definition kept).");
            return;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                notes.Add($"'{funcName}' was NOT created - no valid new name was provided (existing definition kept).");
                return;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                notes.Add($"'{funcName}' was NOT created - the alternate name '{newName}' also already exists.");
                return;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
};

for (int i = 0; i < combos.Count; i++)
{
    var (n, isUpper) = combos[i];
    string funcName    = string.IsNullOrWhiteSpace(nameBoxes[i].Text) ? SuggestName(n, isUpper) : nameBoxes[i].Text.Trim();
    string daxBody     = BuildAxisBody(n, isUpper);
    string description = BuildDescription(n, isUpper);
    CreateUDF(funcName, daxBody, description);
}

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = $"Done! Requested {combos.Count} visual calculation axis UDF(s) for {selectedCounts.Count} input count(s).";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n2 => "- " + n2));

Info(doneSummary);
