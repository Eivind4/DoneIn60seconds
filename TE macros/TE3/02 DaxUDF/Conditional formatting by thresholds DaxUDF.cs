// ADD CONDITIONAL FORMATTING AS DAX UDFs
// ========================================
//
// Purpose: Generate reusable DAX User Defined Functions (UDFs) for conditional formatting of
// visuals (background color / font color, and a plain "good"/"bad"/"neutral" text flag), based on
// a base measure (typically a % difference to a target) compared against an upper and lower bound.
//
// Background vs text:
//   - Background: the functions WITHOUT "_text" return "good" / "bad" / "neutral", which use the
//     report theme's sentiment colors. Nothing to set.
//   - Text (font color): the "_text" functions return a hex color. Default: the theme's
//     "foreground" color for good, bad and neutral (readable on top of the background colors;
//     single measures can get their own color later in the report). The theme is read from the
//     report next to the open model, or from a report / theme file you select. "Choose colors"
//     sets one color per result instead, default white (reads on blue, green, red and orange).
//
// Two model-independent generic functions are (re)created (used as the actual comparison logic):
//   - ConditionalFormat.IncreaseIsGood        -> returns "good" / "bad" / "neutral"
//   - ConditionalFormat.IncreaseIsGood_text    -> returns a hex TEXT color for "good" / "bad" / "neutral"
//   - ConditionalFormat.DecreaseIsGood         -> returns "good" / "bad" / "neutral" (reversed logic)
//   - ConditionalFormat.DecreaseIsGood_text    -> returns a hex TEXT color (reversed logic)
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
using System.IO;
using Newtonsoft.Json.Linq;

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
fixedRadio.Margin = new Padding(3, 8, 3, 0);

// -- Fixed panel --
var fixedPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0) };
fixedPanel.Controls.Add(new Label { Text = "Minimum value (bad below this):", AutoSize = true });
var minFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = -15 };
fixedPanel.Controls.Add(minFixedInput);
fixedPanel.Controls.Add(new Label { Text = "Maximum value (good above this):", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxFixedInput = new NumericUpDown { Width = 120, Minimum = -100000, Maximum = 100000, DecimalPlaces = 2, Increment = 1, Value = 15 };
fixedPanel.Controls.Add(maxFixedInput);

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

// Each option is followed directly by its own inputs: dynamic first, then fixed values
inputFlow.Controls.Add(dynamicRadio);
inputFlow.Controls.Add(dynamicPanel);
inputFlow.Controls.Add(fixedRadio);
inputFlow.Controls.Add(fixedPanel);

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
// DIALOG 2: TEXT (font) color returned by the "_text" functions.
//   - Background formatting uses the functions WITHOUT "_text": they return "good" / "bad" /
//     "neutral", which Power BI maps to the report theme's sentiment colors - nothing to set here.
//   - Default: the theme's "foreground" (text) color for good, bad AND neutral. Change single
//     measures to specific colors later in the report if needed.
//   - Custom: one color per result. Defaults to white, which reads well on blue, green, red and
//     orange backgrounds. Each color is previewed on its own theme background (the theme's good /
//     bad / neutral color, i.e. what the background formatting shows), then on the four samples.
// The theme is read from a report that points to the open model (found automatically when the
// model was opened from a PBIP folder), or from a report / theme file you select.
// ================================================================================================
// Power BI default theme values, used when no report is found or the theme does not set them.
// good / bad / neutral are the sentiment colors the background formatting uses.
var themeDefaults = new Dictionary<string, string> {
    ["foreground"] = "#252423", ["good"] = "#1AAB40", ["bad"] = "#D64554", ["neutral"] = "#D9B300" };
var themeColors = new Dictionary<string, string>(themeDefaults);
string themeForeground = themeColors["foreground"];
string themeName = "Power BI default theme";

// Reads foreground + sentiment colors from a PBIR report: definition.pbir (or the .Report folder)
// -> report.json themeCollection -> base theme, then the custom theme on top (custom wins, as in
// Power BI). A theme .json file can also be picked directly. Returns an error message, or null.
Func<string, string> LoadThemeForeground = path =>
{
    try
    {
        var themeFiles = new List<string>();
        if (path.EndsWith(".json", StringComparison.OrdinalIgnoreCase) && !path.EndsWith("report.json", StringComparison.OrdinalIgnoreCase))
        {
            themeFiles.Add(path);
        }
        else
        {
            string reportFolder = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
            string reportJson = Path.Combine(reportFolder, "definition", "report.json");
            if (!File.Exists(reportJson)) return "No definition\\report.json found below " + reportFolder;
            var themeCollection = JObject.Parse(File.ReadAllText(reportJson))["themeCollection"];
            foreach (var key in new[] { "baseTheme", "customTheme" })
            {
                var theme = themeCollection?[key];
                if (theme == null) continue;
                string name = (string)theme["name"];
                string type = (string)theme["type"];
                string folder = type == "SharedResources"
                    ? Path.Combine(reportFolder, "StaticResources", "SharedResources", "BaseThemes")
                    : Path.Combine(reportFolder, "StaticResources", type ?? "RegisteredResources");
                string file = Path.Combine(folder, name.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? name : name + ".json");
                if (File.Exists(file)) themeFiles.Add(file);
            }
            if (themeFiles.Count == 0) return "The report's theme files were not found - Power BI default theme colors are used.";
        }

        var colors = new Dictionary<string, string>(themeDefaults);
        string loadedName = null;
        foreach (var file in themeFiles)
        {
            var themeJson = JObject.Parse(File.ReadAllText(file));
            loadedName = (string)themeJson["name"] ?? Path.GetFileNameWithoutExtension(file);
            foreach (var key in themeDefaults.Keys)
            {
                string value = (string)themeJson[key];
                if (!string.IsNullOrWhiteSpace(value)) colors[key] = value.Trim();
            }
        }
        themeColors = colors;
        themeForeground = colors["foreground"];
        themeName = loadedName;
        return null;
    }
    catch (Exception ex)
    {
        return "Could not read the theme: " + ex.Message;
    }
};

// Find a report next to the open model whose definition.pbir points to it (PBIP layout)
string themeStatusText = "No report found for the open model - Power BI default theme colors are used. Select a report or theme file to use its theme.";
try
{
    var handler = Model.GetType().GetProperty("Handler")?.GetValue(Model, null);
    var source = handler?.GetType().GetProperty("Source")?.GetValue(handler, null) as string;
    string modelFolder = null;
    for (var d = source == null ? null : (File.Exists(source) ? Path.GetDirectoryName(source) : source); d != null; d = Path.GetDirectoryName(d))
        if (d.EndsWith(".SemanticModel", StringComparison.OrdinalIgnoreCase)) { modelFolder = d; break; }
    if (modelFolder != null)
    {
        foreach (var pbir in Directory.GetFiles(Path.GetDirectoryName(modelFolder), "definition.pbir", SearchOption.AllDirectories))
        {
            var byPath = (string)JObject.Parse(File.ReadAllText(pbir))["datasetReference"]?["byPath"]?["path"];
            if (string.IsNullOrWhiteSpace(byPath)) continue;
            var target = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(pbir), byPath)).TrimEnd('\\', '/');
            if (!string.Equals(target, modelFolder.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase)) continue;
            var error = LoadThemeForeground(pbir);
            themeStatusText = error ?? "Theme read automatically from the report " + Path.GetFileName(Path.GetDirectoryName(pbir)) + ".";
            break;
        }
    }
}
catch { }

// Sample backgrounds the text must be readable on
var sampleBackgrounds = new[] {
    ("blue", ColorTranslator.FromHtml("#118DFF")), ("green", ColorTranslator.FromHtml("#1AAB40")),
    ("red", ColorTranslator.FromHtml("#D64554")), ("orange", ColorTranslator.FromHtml("#E66C37")) };
const string defaultCustomText = "#FFFFFF";   // white: readable on all four sample backgrounds

var colorForm = new Form
{
    Text = "Conditional formatting - TEXT color (_text functions)",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 560,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var colorFlow = new FlowLayoutPanel { Padding = new Padding(12), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
colorForm.Controls.Add(colorFlow);
colorFlow.Controls.Add(new Label {
    Text = "TEXT FORMATTING - font color",
    AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.White, BackColor = Color.FromArgb(17, 141, 255),
    Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 0, 6) });
colorFlow.Controls.Add(new Label {
    Text = "This sets the hex color returned by the '_text' functions, used as font color.\n" +
           "Background formatting needs no setting: the functions without '_text' return good / bad / neutral, " +
           "which use the report theme's sentiment colors.",
    AutoSize = true, MaximumSize = new Size(520, 0), Margin = new Padding(0, 0, 0, 10) });

// Parses a hex color; falls back when the theme holds something unexpected
Func<string, Color, Color> ToColor = (hex, fallback) => { try { return ColorTranslator.FromHtml(hex); } catch { return fallback; } };

// Preview chips with the theme's background (sentiment) colors - refreshed when another theme is
// selected. key = good / bad / neutral; textKey = null for "use the row's own text color".
var themeChips = new List<(Label chip, string key)>();
Func<string, string, int, Label> ThemeChip = (key, text, width) =>
{
    var chip = new Label { Text = text, Width = width, Height = 22, TextAlign = ContentAlignment.MiddleCenter,
        Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 3, 4, 0) };
    themeChips.Add((chip, key));
    return chip;
};

// -- Theme: read automatically, or select another report / theme file (used by both options) --
var themeRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
var themeNameLabel = new Label { AutoSize = true, Margin = new Padding(0, 6, 8, 0) };
var themeBrowse = new Button { Text = "Other report or theme...", AutoSize = true };
themeRow.Controls.Add(themeNameLabel);
themeRow.Controls.Add(themeBrowse);
colorFlow.Controls.Add(themeRow);
var themeStatus = new Label { Text = themeStatusText, AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(520, 0), Margin = new Padding(3, 0, 0, 10) };
colorFlow.Controls.Add(themeStatus);

// -- Theme foreground for all (default) --
var themeRadio = new RadioButton { Text = "Theme foreground color for good, bad and neutral (recommended)", AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
colorFlow.Controls.Add(themeRadio);
var themePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Margin = new Padding(20, 0, 0, 8) };
var themePreviewRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
themePreviewRow.Controls.Add(new Label { Text = "Preview on the theme backgrounds:", AutoSize = true, Margin = new Padding(0, 6, 8, 0) });
themePreviewRow.Controls.Add(ThemeChip("good", "Good", 64));
themePreviewRow.Controls.Add(ThemeChip("bad", "Bad", 64));
themePreviewRow.Controls.Add(ThemeChip("neutral", "Neutral", 64));
themePanel.Controls.Add(themePreviewRow);
themePanel.Controls.Add(new Label { Text = "Specific measures can get their own color later in the report.", AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(500, 0) });
colorFlow.Controls.Add(themePanel);

// -- Custom colors --
var customRadio = new RadioButton { Text = "Choose colors (one per result)", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
colorFlow.Controls.Add(customRadio);
var customPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Margin = new Padding(20, 0, 0, 0) };
customPanel.Controls.Add(new Label {
    Text = "Each color is previewed on its own theme background (good / bad / neutral), then on blue, green, red and orange. Default white reads well on all of them.",
    AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(520, 0), Margin = new Padding(0, 0, 0, 6) });
colorFlow.Controls.Add(customPanel);

// One "Label | hex textbox | Pick... | preview on its theme background | on blue/green/red/orange"
// row; the previews follow the textbox whether the user types a hex code or uses the color picker.
var customTextColors = new Dictionary<Label, TextBox>();
Func<string, string, TextBox> AddColorRow = (rowLabel, themeKey) =>
{
    var row = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
    row.Controls.Add(new Label { Text = rowLabel, Width = 60, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 0, 0) });
    var box = new TextBox { Width = 80, Text = defaultCustomText, Margin = new Padding(0, 3, 6, 0) };
    var pickBtn = new Button { Text = "Pick...", AutoSize = true, Margin = new Padding(0, 1, 10, 0) };
    var themeChip = ThemeChip(themeKey, "Theme " + themeKey, 96);
    customTextColors[themeChip] = box;
    var previews = sampleBackgrounds.Select(bg => new Label {
        Text = "Text", Width = 44, Height = 22, TextAlign = ContentAlignment.MiddleCenter, BackColor = bg.Item2,
        Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 3, 2, 0) }).ToList();
    previews.Insert(0, themeChip);

    Action refresh = () =>
    {
        try { var c = ColorTranslator.FromHtml(box.Text.Trim()); foreach (var p in previews) p.ForeColor = c; }
        catch { /* leave the preview unchanged until a valid hex is entered */ }
    };
    box.TextChanged += (s, e) => refresh();
    pickBtn.Click += (s, e) =>
    {
        using (var cd = new ColorDialog { FullOpen = true })
        {
            try { cd.Color = ColorTranslator.FromHtml(box.Text.Trim()); } catch { }
            if (cd.ShowDialog() == DialogResult.OK)
                box.Text = "#" + cd.Color.R.ToString("X2") + cd.Color.G.ToString("X2") + cd.Color.B.ToString("X2");
        }
    };
    refresh();

    row.Controls.Add(box);
    row.Controls.Add(pickBtn);
    foreach (var p in previews) row.Controls.Add(p);
    customPanel.Controls.Add(row);
    return box;
};

var goodBox    = AddColorRow("Good:", "good");
var badBox     = AddColorRow("Bad:", "bad");
var neutralBox = AddColorRow("Neutral:", "neutral");

Action applyColorMode = () =>
{
    themePanel.Enabled = themeRadio.Checked;
    customPanel.Enabled = customRadio.Checked;
    themeNameLabel.Text = "Theme: " + themeName + "   foreground " + themeForeground +
        "   good " + themeColors["good"] + "   bad " + themeColors["bad"] + "   neutral " + themeColors["neutral"];
    // Theme backgrounds on every chip; text = theme foreground, or the custom row's own color
    var foreground = ToColor(themeForeground, Color.Black);
    foreach (var (chip, key) in themeChips)
    {
        chip.BackColor = ToColor(themeColors[key], Color.Gray);
        TextBox box;
        chip.ForeColor = customTextColors.TryGetValue(chip, out box) ? ToColor(box.Text.Trim(), chip.ForeColor) : foreground;
    }
};
themeRadio.CheckedChanged += (s, e) => applyColorMode();
customRadio.CheckedChanged += (s, e) => applyColorMode();
themeBrowse.Click += (s, e) =>
{
    using (var ofd = new OpenFileDialog { Title = "Select the report (definition.pbir) or a theme .json file",
        Filter = "Report definition or theme (definition.pbir;*.json)|definition.pbir;*.json|All files (*.*)|*.*" })
    {
        if (ofd.ShowDialog() != DialogResult.OK) return;
        string error = LoadThemeForeground(ofd.FileName);
        themeStatus.Text = error ?? "Theme read from " + Path.GetFileName(ofd.FileName) + ".";
        themeStatus.ForeColor = error == null ? Color.DimGray : Color.Firebrick;
        applyColorMode();
    }
};
applyColorMode();

var colorBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 12, 0, 0) };
var colorOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var colorCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
colorBtnPanel.Controls.Add(colorOk);
colorBtnPanel.Controls.Add(colorCancel);
colorFlow.Controls.Add(colorBtnPanel);
colorForm.AcceptButton = colorOk;
colorForm.CancelButton = colorCancel;

if (colorForm.ShowDialog() == DialogResult.Cancel) return;

bool useThemeForeground = themeRadio.Checked;
string goodColor    = useThemeForeground ? themeForeground : goodBox.Text.Trim();
string badColor     = useThemeForeground ? themeForeground : badBox.Text.Trim();
string neutralColor = useThemeForeground ? themeForeground : neutralBox.Text.Trim();
notes.Add(useThemeForeground
    ? $"Text color: theme '{themeName}' foreground {themeForeground} for good, bad and neutral."
    : $"Text colors: good {goodColor}, bad {badColor}, neutral {neutralColor}.");

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
    "Returns a hex text (font) color - an increase (value above the upper bound) is good.");
string decreaseTextName = CreateGenericUDF("ConditionalFormat.DecreaseIsGood_text", decreaseTextBody,
    "Returns a hex text (font) color - a decrease (value below the lower bound) is good.");

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
    "Text (font) color for conditional formatting where an increase is good.", textPrefix, folderSuffix);
CreateLocalUDF($"{localPrefix}.DecreaseIsGood_text", LocalWrapperBody(decreaseTextName),
    "Text (font) color for conditional formatting where a decrease is good.", textPrefix, folderSuffix);

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = "Done! Created/updated the ConditionalFormat generic functions and their Local wrapper functions.";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n => "- " + n));

Info(doneSummary);
