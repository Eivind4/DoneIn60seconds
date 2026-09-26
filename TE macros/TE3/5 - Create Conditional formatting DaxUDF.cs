// ADD CONDITIONAL FORMATTING AS DAX UDFs
// ========================================
//
// Purpose: Generate reusable DAX User Defined Functions (UDFs) for conditional formatting of
// visuals (background color / font color, and a plain "good"/"bad"/"neutral" text flag), based on
// a base measure (typically a % difference to a target) compared against an upper and lower bound.
//
// Two model-independent generic functions are (re)created (used as the actual comparison logic):
//   - ConditionalFormat.IncreaseIsGood        -> returns "good" / "bad" / "neutral"
//   - ConditionalFormat.IncreaseIsGood_text    -> returns a hex color for "good" / "bad" / "neutral"
//   - ConditionalFormat.DecreaseIsGood         -> returns "good" / "bad" / "neutral" (reversed logic)
//   - ConditionalFormat.DecreaseIsGood_text    -> returns a hex color (reversed logic)
//
// Four model-DEPENDENT "Local" wrapper functions are also created, each taking only a single
// baseMeasure parameter and supplying the upper/lower bound (either fixed literal values, or
// measures from a "Conditional Format Input" calculated table you configure below):
//   - Local.ConditionalFormat.IncreaseIsGood        (+ "_text" variant)
//   - Local.ConditionalFormat.DecreaseIsGood         (+ "_text" variant)
// These get annotations (FolderSuffix / MeasurePrefix) so they can be turned into measures using
// the "5 - Create measure from model dependent functions.cs" macro.
//
// Safety checks performed before creating each function (generic and Local):
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function is reused/referenced instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one (and, for Local
//       wrappers, adjust the MeasurePrefix annotation), Overwrite the existing (outdated)
//       definition, or Skip creating it (never silently overwritten).
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
// DIALOG 1: Fixed target input vs Dynamic (calculated table), plus the min/max/increment values
// and whether they represent percentage points (e.g. entering 15 for 15%).
// ================================================================================================
var inputForm = new Form
{
    Text = "Conditional formatting - input source",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 480,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var inputFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
inputForm.Controls.Add(inputFlow);
Action<string> inputLabel = s => inputFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

inputLabel("Choose how the upper/lower bound is determined:");
var fixedRadio   = new RadioButton { Text = "Fixed target input (hard-coded min/max values)", AutoSize = true, Checked = true };
var dynamicRadio = new RadioButton { Text = "Dynamic (adds a 'Conditional Format Input' calculated table with a slicer)", AutoSize = true };
inputFlow.Controls.Add(fixedRadio);
inputFlow.Controls.Add(dynamicRadio);

// -- Fixed panel --
var fixedPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0) };
fixedPanel.Controls.Add(new Label { Text = "Minimum value (bad below this):", AutoSize = true });
var minFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = -15 };
fixedPanel.Controls.Add(minFixedInput);
fixedPanel.Controls.Add(new Label { Text = "Maximum value (good above this):", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = 15 };
fixedPanel.Controls.Add(maxFixedInput);
inputFlow.Controls.Add(fixedPanel);

// -- Dynamic panel --
var dynamicPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0), Visible = false };
dynamicPanel.Controls.Add(new Label { Text = "Table minimum:", AutoSize = true });
var minDynInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, Value = -30 };
dynamicPanel.Controls.Add(minDynInput);
dynamicPanel.Controls.Add(new Label { Text = "Table maximum:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxDynInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, Value = 30 };
dynamicPanel.Controls.Add(maxDynInput);
dynamicPanel.Controls.Add(new Label { Text = "Increment:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var incDynInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 10000, Value = 1 };
dynamicPanel.Controls.Add(incDynInput);
inputFlow.Controls.Add(dynamicPanel);

var percentCheckbox = new CheckBox { Text = "Values represent percentage points (e.g. enter 15 for 15% - divided by 100)", AutoSize = true, Checked = true, Margin = new Padding(0, 10, 0, 0) };
inputFlow.Controls.Add(percentCheckbox);

Action updateInputVis = () => { fixedPanel.Visible = fixedRadio.Checked; dynamicPanel.Visible = dynamicRadio.Checked; };
fixedRadio.CheckedChanged   += (s, e) => updateInputVis();
dynamicRadio.CheckedChanged += (s, e) => updateInputVis();

var inputBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var inputOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var inputCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
inputBtnPanel.Controls.Add(inputOk);
inputBtnPanel.Controls.Add(inputCancel);
inputFlow.Controls.Add(inputBtnPanel);
inputForm.AcceptButton = inputOk;
inputForm.CancelButton = inputCancel;

if (inputForm.ShowDialog() == DialogResult.Cancel) return;

bool isDynamic    = dynamicRadio.Checked;
bool isPercentage = percentCheckbox.Checked;

// ================================================================================================
// Resolve the upper/lower bound DAX references used inside the Local wrapper functions - either
// literal fixed values, or references to measures on a "Conditional Format Input" calculated table.
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
else
{
    const string inputTableName = "Conditional Format Input";
    var existingTable = Model.Tables.FirstOrDefault(t => t.Name == inputTableName);

    string maxMeasureName = isPercentage ? "Max Value percent" : "Max Value";
    string minMeasureName = isPercentage ? "Min Value percent" : "Min Value";

    bool createTableAndMeasures = true;

    if (existingTable != null)
    {
        var choice = MessageBox.Show(
            $"A table named '{inputTableName}' already exists.\n\n" +
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

        var table = Model.AddCalculatedTable(inputTableName, $"GENERATESERIES({min}, {max}, {inc})");
        table.Description = "To set the dynamic max and min value for conditional formatting";

        // TE3 auto-creates the "Value" column from GENERATESERIES; fall back to the first column
        // in case a differently-named/typed column is inferred.
        var valueColumn = table.Columns.Contains("Value") ? table.Columns["Value"] : table.Columns.First();
        string valueColRef = valueColumn.DaxObjectFullName;
        string tableRef    = table.DaxObjectFullName;

        var valueMeasure = table.AddMeasure(
            $"{inputTableName} Value",
            $"SELECTEDVALUE({valueColRef}, 0)",
            "Conditional Formatting");
        valueMeasure.Description = $"Selected value from the {inputTableName} table, defaults to 0 if nothing selected";
        valueMeasure.FormatString = "#,0";

        string divideSuffix = isPercentage ? " / 100" : "";
        string formatString = isPercentage ? "#,0.0%;-#,0.0%;#,0.0%" : "#,0";

        var maxMeasure = table.AddMeasure(
            maxMeasureName,
            $"MAXX({tableRef}, {valueColRef}){divideSuffix}",
            "Conditional Formatting");
        maxMeasure.Description = isPercentage
            ? $"Maximum value from {inputTableName} table divided by 100 to convert to decimal"
            : $"Maximum value from {inputTableName} table";
        maxMeasure.FormatString = formatString;

        var minMeasure = table.AddMeasure(
            minMeasureName,
            $"MINX({tableRef}, {valueColRef}){divideSuffix}",
            "Conditional Formatting");
        minMeasure.Description = isPercentage
            ? $"Minimum value from {inputTableName} table divided by 100 to convert to decimal"
            : $"Minimum value from {inputTableName} table";
        minMeasure.FormatString = formatString;

        existingTable = table;
    }

    var maxMeasureRef = existingTable.Measures.FirstOrDefault(m => m.Name == maxMeasureName);
    var minMeasureRef = existingTable.Measures.FirstOrDefault(m => m.Name == minMeasureName);
    if (maxMeasureRef == null || minMeasureRef == null)
    {
        Error($"Could not find measures '{maxMeasureName}' / '{minMeasureName}' on the existing '{inputTableName}' table. Re-run and choose Overwrite, or align the percentage option with how the table was originally created.");
        return;
    }

    upperArg = maxMeasureRef.DaxObjectFullName;
    lowerArg = minMeasureRef.DaxObjectFullName;
}

// ================================================================================================
// DIALOG 2: Colors used by the "_text" (hex color) functions, with a live preview swatch.
// ================================================================================================
string goodColor = "#FFFFFF";
string badColor = "#FFFFFF";
string neutralColor = "#183A46";

var colorForm = new Form
{
    Text = "Conditional formatting - colors",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 460,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var colorFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
colorForm.Controls.Add(colorFlow);
colorFlow.Controls.Add(new Label { Text = "Choose the hex colors returned by the '_text' functions:", AutoSize = true, Margin = new Padding(0, 0, 0, 10) });

// Builds one "Label | hex textbox | Pick... button | preview swatch" row, wired so the swatch and
// textbox stay in sync whether the user types a hex code or uses the color picker.
Func<string, string, (TextBox box, Panel swatch)> AddColorRow = (rowLabel, defaultHex) =>
{
    var row = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
    row.Controls.Add(new Label { Text = rowLabel, Width = 70, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 0, 0) });
    var box = new TextBox { Width = 90, Text = defaultHex, Margin = new Padding(0, 3, 8, 0) };
    var swatch = new Panel { Width = 30, Height = 22, BorderStyle = BorderStyle.FixedSingle, BackColor = ColorTranslator.FromHtml(defaultHex), Margin = new Padding(0, 3, 8, 0) };
    var pickBtn = new Button { Text = "Pick...", AutoSize = true, Margin = new Padding(0, 1, 0, 0) };

    Action refreshSwatch = () =>
    {
        try { swatch.BackColor = ColorTranslator.FromHtml(box.Text.Trim()); }
        catch { /* leave swatch unchanged until a valid hex is entered */ }
    };
    box.TextChanged += (s, e) => refreshSwatch();
    pickBtn.Click += (s, e) =>
    {
        using (var cd = new ColorDialog { Color = swatch.BackColor, FullOpen = true })
        {
            if (cd.ShowDialog() == DialogResult.OK)
            {
                box.Text = "#" + cd.Color.R.ToString("X2") + cd.Color.G.ToString("X2") + cd.Color.B.ToString("X2");
            }
        }
    };

    row.Controls.Add(box);
    row.Controls.Add(swatch);
    row.Controls.Add(pickBtn);
    colorFlow.Controls.Add(row);
    return (box, swatch);
};

var (goodBox, goodSwatch)       = AddColorRow("Good:", goodColor);
var (badBox, badSwatch)         = AddColorRow("Bad:", badColor);
var (neutralBox, neutralSwatch) = AddColorRow("Neutral:", neutralColor);

var colorBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
var colorOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var colorCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
colorBtnPanel.Controls.Add(colorOk);
colorBtnPanel.Controls.Add(colorCancel);
colorFlow.Controls.Add(colorBtnPanel);
colorForm.AcceptButton = colorOk;
colorForm.CancelButton = colorCancel;

if (colorForm.ShowDialog() == DialogResult.Cancel) return;

goodColor    = goodBox.Text.Trim();
badColor     = badBox.Text.Trim();
neutralColor = neutralBox.Text.Trim();

// ================================================================================================
// Create the 4 model-independent generic functions (the actual comparison logic).
// Returns the effective name that should be referenced afterwards (may differ from funcName if
// an identical expression already existed elsewhere, or the user chose to Rename on conflict).
// ================================================================================================
Func<string, string, string, string> CreateGenericUDF = (funcName, daxBody, description) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        notes.Add($"'{funcName}' was NOT created - using the existing function '{dupExprFunc.Name}' instead, since it has an identical expression.");
        return dupExprFunc.Name;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            notes.Add($"'{funcName}' already exists with the same expression - left unchanged.");
            return existing.Name;
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
            return existing.Name;
        }
        else if (skipRadio.Checked)
        {
            notes.Add($"'{funcName}' was NOT created - a function with that name already exists (existing definition kept).");
            return existing.Name;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                notes.Add($"'{funcName}' was NOT created - no valid new name was provided (existing definition kept).");
                return existing.Name;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                notes.Add($"'{funcName}' was NOT created - the alternate name '{newName}' also already exists.");
                return existing.Name;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
    return fn.Name;
};

string increaseColorBody = $@"(
    // What you are measuring - typically a % difference to a target
    baseMeasure: ANYREF,
    // A measure or fixed number giving the upper bound - values above this are ""good""
    upperPercent: ANYREF,
    // A measure or fixed number giving the lower bound - values below this are ""bad""
    lowerPercent: ANYREF
) =>
    VAR _TargetPercent = baseMeasure
    VAR _PctChangePos = upperPercent
    VAR _PctChangeNeg = lowerPercent
    RETURN
        SWITCH(
            TRUE(),
            _TargetPercent > _PctChangePos, ""good"",
            _TargetPercent < _PctChangeNeg, ""bad"",
            ""neutral""
        )";

string decreaseColorBody = $@"(
    // What you are measuring - typically a % difference to a target
    baseMeasure: ANYREF,
    // A measure or fixed number giving the upper bound - values above this are ""bad""
    upperPercent: ANYREF,
    // A measure or fixed number giving the lower bound - values below this are ""good""
    lowerPercent: ANYREF
) =>
    VAR _TargetPercent = baseMeasure
    VAR _PctChangePos = upperPercent
    VAR _PctChangeNeg = lowerPercent
    RETURN
        SWITCH(
            TRUE(),
            _TargetPercent < _PctChangeNeg, ""good"",
            _TargetPercent > _PctChangePos, ""bad"",
            ""neutral""
        )";

string increaseTextBody = $@"(
    // What you are measuring - typically a % difference to a target
    baseMeasure: ANYREF,
    // A measure or fixed number giving the upper bound - values above this are ""good""
    upperPercent: ANYREF,
    // A measure or fixed number giving the lower bound - values below this are ""bad""
    lowerPercent: ANYREF
) =>
    VAR _TargetPercent = baseMeasure
    VAR _PctChangePos = upperPercent
    VAR _PctChangeNeg = lowerPercent
    RETURN
        SWITCH(
            TRUE(),
            _TargetPercent > _PctChangePos, ""{goodColor}"",
            _TargetPercent < _PctChangeNeg, ""{badColor}"",
            ""{neutralColor}""
        )";

string decreaseTextBody = $@"(
    // What you are measuring - typically a % difference to a target
    baseMeasure: ANYREF,
    // A measure or fixed number giving the upper bound - values above this are ""bad""
    upperPercent: ANYREF,
    // A measure or fixed number giving the lower bound - values below this are ""good""
    lowerPercent: ANYREF
) =>
    VAR _TargetPercent = baseMeasure
    VAR _PctChangePos = upperPercent
    VAR _PctChangeNeg = lowerPercent
    RETURN
        SWITCH(
            TRUE(),
            _TargetPercent < _PctChangeNeg, ""{goodColor}"",
            _TargetPercent > _PctChangePos, ""{badColor}"",
            ""{neutralColor}""
        )";

string increaseColorName = CreateGenericUDF("ConditionalFormat.IncreaseIsGood", increaseColorBody,
    "Returns 'good'/'bad'/'neutral' - an increase (value above the upper bound) is good.");
string decreaseColorName = CreateGenericUDF("ConditionalFormat.DecreaseIsGood", decreaseColorBody,
    "Returns 'good'/'bad'/'neutral' - a decrease (value below the lower bound) is good.");
string increaseTextName = CreateGenericUDF("ConditionalFormat.IncreaseIsGood_text", increaseTextBody,
    "Returns a hex color - an increase (value above the upper bound) is good.");
string decreaseTextName = CreateGenericUDF("ConditionalFormat.DecreaseIsGood_text", decreaseTextBody,
    "Returns a hex color - a decrease (value below the lower bound) is good.");

// ================================================================================================
// DIALOG 3: Local wrapper naming - function name prefix, FolderSuffix annotation, and MeasurePrefix
// annotations (used by "5 - Create measure from model dependent functions.cs" to name/organize the
// generated measures). The prefix does not differentiate Increase/Decrease, only color vs. text.
// ================================================================================================
var localForm = new Form
{
    Text = "Conditional formatting - Local wrapper naming",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 460,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var localFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
localForm.Controls.Add(localFlow);
Action<string> localLabel = s => localFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

localLabel("(Optional) Function name prefix  [default: Local.ConditionalFormat]:");
var localPrefixTextbox = new TextBox { Width = 350, Text = "Local.ConditionalFormat" };
localFlow.Controls.Add(localPrefixTextbox);

localLabel("(Optional) FolderSuffix annotation value  [default: Conditional Formatting]:");
var folderSuffixTextbox = new TextBox { Width = 350, Text = "Conditional Formatting" };
localFlow.Controls.Add(folderSuffixTextbox);

localLabel("(Optional) MeasurePrefix annotation for the color functions  [default: CF]:");
var colorPrefixTextbox = new TextBox { Width = 350, Text = "CF" };
localFlow.Controls.Add(colorPrefixTextbox);

localLabel("(Optional) MeasurePrefix annotation for the text functions  [default: CF text]:");
var textPrefixTextbox = new TextBox { Width = 350, Text = "CF text" };
localFlow.Controls.Add(textPrefixTextbox);

var localBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var localOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var localCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
localBtnPanel.Controls.Add(localOk);
localBtnPanel.Controls.Add(localCancel);
localFlow.Controls.Add(localBtnPanel);
localForm.AcceptButton = localOk;
localForm.CancelButton = localCancel;

if (localForm.ShowDialog() == DialogResult.Cancel) return;

string localPrefix   = string.IsNullOrWhiteSpace(localPrefixTextbox.Text) ? "Local.ConditionalFormat" : localPrefixTextbox.Text.Trim();
string folderSuffix  = folderSuffixTextbox.Text.Trim();
string colorPrefix   = colorPrefixTextbox.Text.Trim();
string textPrefix    = textPrefixTextbox.Text.Trim();

// ================================================================================================
// Create the 4 model-dependent Local wrapper functions, each calling the matching generic function
// above with either the fixed literal bounds or the calculated-table measure references.
// ================================================================================================
Action<string, string, string, string, string> CreateLocalUDF = (funcName, daxBody, description, measurePrefix, defaultFolderSuffix) =>
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

        var conflictDlg = new Form { Text = "Function name already exists", Width = 500, Height = 300, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false };
        var conflictFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = $"A function named '{funcName}' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(460, 0), Margin = new Padding(0, 0, 0, 10) });

        var renameRow   = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var renameRadio = new RadioButton { Text = "Rename to:", AutoSize = true, Checked = true, Margin = new Padding(0, 3, 8, 0) };
        var renameBox   = new TextBox { Width = 250, Text = funcName };
        renameRow.Controls.Add(renameRadio);
        renameRow.Controls.Add(renameBox);
        conflictFlow.Controls.Add(renameRow);

        var prefixRow  = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(20, 0, 0, 5) };
        prefixRow.Controls.Add(new Label { Text = "MeasurePrefix annotation to use:", AutoSize = true, Margin = new Padding(0, 3, 8, 0) });
        var prefixBox = new TextBox { Width = 150, Text = measurePrefix };
        prefixRow.Controls.Add(prefixBox);
        conflictFlow.Controls.Add(prefixRow);

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
        measurePrefix = string.IsNullOrWhiteSpace(prefixBox.Text) ? measurePrefix : prefixBox.Text.Trim();

        if (overwriteRadio.Checked)
        {
            existing.Expression  = daxBody;
            existing.Description = description;
            existing.SetAnnotation("FolderSuffix",  defaultFolderSuffix);
            existing.SetAnnotation("MeasurePrefix", measurePrefix);
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
    fn.SetAnnotation("FolderSuffix",  defaultFolderSuffix);
    fn.SetAnnotation("MeasurePrefix", measurePrefix);
};

Func<string, string> LocalWrapperBody = genericFuncName => $@"(
    // A base measure, typically returning the % difference to a target
    baseMeasure: ANYREF
) =>
    {genericFuncName}(
        baseMeasure,
        {upperArg},
        {lowerArg}
    )";

CreateLocalUDF($"{localPrefix}.IncreaseIsGood", LocalWrapperBody(increaseColorName),
    "Applies conditional formatting ('good'/'bad'/'neutral') where an increase is good.", colorPrefix, folderSuffix);
CreateLocalUDF($"{localPrefix}.DecreaseIsGood", LocalWrapperBody(decreaseColorName),
    "Applies conditional formatting ('good'/'bad'/'neutral') where a decrease is good.", colorPrefix, folderSuffix);
CreateLocalUDF($"{localPrefix}.IncreaseIsGood_text", LocalWrapperBody(increaseTextName),
    "Applies conditional formatting (hex color) where an increase is good.", textPrefix, folderSuffix);
CreateLocalUDF($"{localPrefix}.DecreaseIsGood_text", LocalWrapperBody(decreaseTextName),
    "Applies conditional formatting (hex color) where a decrease is good.", textPrefix, folderSuffix);

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = "Done! Created/updated the ConditionalFormat generic functions and their Local wrapper functions.";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n => "- " + n));

Info(doneSummary);
