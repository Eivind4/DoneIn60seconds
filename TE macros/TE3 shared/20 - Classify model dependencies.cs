// ============================================================================
// 16 - Classify model dependencies from PBIR.cs
//
// Tabular Editor 3 macro that scans one or more PBIR reports and annotates
// measures/columns in the open semantic model with:
//   DependencyCategory = VisualDependency
//   DependencyCategory = MeasureWithVisualDependency
//   DependencyCategory = DaxUdfWithVisualDependency
//   DependencyCategory = DaxUdfDependencyWithoutVisualDependency
//   DependencyCategory = MeasureDependencyWithoutVisualDependency
//   DependencyCategory = ColumnUsedByDaxUdfDependency
//   DependencyCategory = ColumnUsedByVisualDependentMeasure
//   DependencyCategory = ColumnUsedByMeasureDependencyWithoutVisualDependency
//   DependencyCategory = ModelDependencyWithoutVisualDependency
//   DependencyCategory = NoDependencies
//
// Workflow:
//   1. Pick a folder. The macro finds every definition.pbir below it.
//   2. Select which discovered reports to include.
//   3. Choose whether to classify measures, columns, or both. If measures or
//      columns are selected in Tabular Editor, you can scope to selected objects.
//   4. PBIR visual/report/page JSON is scanned for structural Measure/Column
//      references, including nested filter expressions and SourceRef aliases.
//   5. Measures are classified by direct visual use and upstream model
//      dependencies. Columns are classified by direct visual use, relationship /
//      hierarchy / sort-by usage, and measure dependencies.
//   6. Results are shown in the output grid (filterable), including a
//      "Duplicate / redundant" column: measures with identical DAX, columns
//      with the same name in other tables, and columns with the same DAX or
//      source column as another column in the same table.
//   7. One dialog with horizontal buttons:
//        Move to 'XX - Not to be used' | Choose measures to delete |
//        Keep annotations only | Do nothing
//      Every option except Do nothing writes the DependencyCategory annotations.
//
// 'XX - Not to be used' folder:
//   - Not-used measures are moved to this ROOT display folder (one per table),
//     so they are easy to find. The original folder is kept in the annotation
//     DependencyOriginalDisplayFolder.
//   - When the macro runs again and such a measure is used again (e.g. a visual
//     now references it), it is moved back to its original folder. Measures in
//     the older nested 'Not used in report' folder are handled the same way.
//
// Notes / limits for this first version:
//   - Requires PBIR files to be available on disk.
//   - The open Tabular Editor model is the source of truth for model
//     dependencies. Live-connected report metadata is listed when found in
//     definition.pbir, but service-side model dependencies are not fetched.
//   - Test on a copy/dev model first. Measure deletion is undoable in TE UI,
//     but PBIR scanning itself is read-only.
//   - Only structural Measure/Column references are detected. References that
//     exist only as DAX text in the report layer are not (see ToDo). These
//     gaps are listed in the scope warning shown at start.
//
// ToDo:
//   - reportExtensions.json (report-level measures in thin reports): scan each
//     entity's measure "expression" DAX text for model measures [Name],
//     columns 'Table'[Column] and DAX UDF calls FunctionName( and treat hits as
//     visual dependencies. Most relevant for DAX UDFs, which report measures
//     can call directly.
//   - Visual calculations (NativeVisualCalculation in visual.json) have the
//     same DAX-text gap - e.g. a visual calculation calling a DAX UDF.
//   - bookmarks/*.json: low priority. Bookmarks store filter state, so a column
//     used ONLY as a filter inside a bookmark would be missed. Visuals shown or
//     hidden by bookmarks are still normal visual.json files and are detected.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

const string AnnotationName = "DependencyCategory";
const string CatVisual = "VisualDependency";
const string CatMeasureWithVisual = "MeasureWithVisualDependency";
const string CatDaxUdfWithVisual = "DaxUdfWithVisualDependency";
const string CatDaxUdfNoVisual = "DaxUdfDependencyWithoutVisualDependency";
const string CatMeasureNoVisual = "MeasureDependencyWithoutVisualDependency";
const string CatColumnUsedByVisualMeasure = "ColumnUsedByVisualDependentMeasure";
const string CatColumnUsedByDaxUdfMeasure = "ColumnUsedByDaxUdfDependency";
const string CatColumnUsedByNoVisualMeasure = "ColumnUsedByMeasureDependencyWithoutVisualDependency";
const string CatModelNoVisual = "ModelDependencyWithoutVisualDependency";
const string CatNone = "NoDependencies";
const string UnusedFolder = "XX - Not to be used";          // root display folder for not-used measures
const string LegacyUnusedFolder = "Not used in report";     // nested folder used by older versions
const string OriginalFolderAnnotation = "DependencyOriginalDisplayFolder";

class ReportChoice
{
    public string DefinitionFile;
    public string ReportFolder;
    public string ReportName;
    public string ConnectionSummary;
    public bool Selected;
    public string Label
    {
        get { return ReportName + "  |  " + ConnectionSummary + "  |  " + ReportFolder; }
    }
}

class PbirReference
{
    public string Kind;
    public string Table;
    public string Field;
    public string ReportName;
    public string File;
    public string Key
    {
        get { return Kind + "|" + Table + "|" + Field; }
    }
}

class ObjectResult
{
    public string Kind;
    public string Table;
    public string Name;
    public string Category;
    public string Reason;
    public bool Hidden;
    public int VisualReferences;
    public int DependsOnCount;
    public int ReferencedByCount;
    public Measure MeasureObject;
    public Column ColumnObject;
    public Function FunctionObject;
}

Func<Size> BigDialogSize = () =>
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * 0.85), (int)(wa.Height * 0.85));
};

Func<string, string> CleanReportName = folder =>
{
    var name = Path.GetFileName(folder);
    return name != null && name.EndsWith(".Report", StringComparison.OrdinalIgnoreCase)
        ? name.Substring(0, name.Length - ".Report".Length)
        : name;
};

Func<string, string> ReadConnectionSummary = definitionFile =>
{
    try
    {
        var root = JObject.Parse(File.ReadAllText(definitionFile));
        var dataset = root["datasetReference"] ?? root["semanticModelReference"];
        var byPath = (string)dataset?["byPath"]?["path"];
        if (!string.IsNullOrWhiteSpace(byPath)) return "byPath: " + byPath;
        var conn = (string)dataset?["byConnection"]?["connectionString"];
        if (!string.IsNullOrWhiteSpace(conn)) return "byConnection/live";
    }
    catch { }
    return "connection unknown";
};

Func<object, IEnumerable<object>> GetDependsOn = obj =>
{
    if (obj == null) return Enumerable.Empty<object>();
    var prop = obj.GetType().GetProperty("DependsOn");
    if (prop == null) return Enumerable.Empty<object>();
    var value = prop.GetValue(obj, null) as System.Collections.IEnumerable;
    return value == null ? Enumerable.Empty<object>() : value.Cast<object>();
};

Func<object, IEnumerable<object>> GetReferencedBy = obj =>
{
    if (obj == null) return Enumerable.Empty<object>();
    var prop = obj.GetType().GetProperty("ReferencedBy");
    if (prop == null) return Enumerable.Empty<object>();
    var value = prop.GetValue(obj, null) as System.Collections.IEnumerable;
    return value == null ? Enumerable.Empty<object>() : value.Cast<object>();
};

Func<Table, bool> IsExcludedSupportTable = t =>
{
    if (t == null) return false;
    if (t is CalculationGroupTable) return true;
    try
    {
        if (t.Columns.Any(c => c.GroupByColumns.Any()))
            return true;
    }
    catch { }
    return false;
};

// ---------------------------------------------------------------------------
// Dialog: scope warning shown at start
// ---------------------------------------------------------------------------
using (var form = new Form())
{
    form.Text = "Before you start - scope of this analysis";
    form.AutoSize = true;
    form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.MaximizeBox = false;
    form.MinimizeBox = false;
    form.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);
    var boldFont = new Font("Segoe UI", 9, FontStyle.Bold);
    const int textWidth = 640;

    var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, Padding = new Padding(14) };
    Action<string, Font, Color, Padding> addText = (text, f, color, margin) =>
        layout.Controls.Add(new Label { Text = text, AutoSize = true, MaximumSize = new Size(textWidth, 0), Font = f, ForeColor = color, Margin = margin });

    addText("This analysis only sees the model open in Tabular Editor and the PBIR reports in the folder you select.",
        new Font("Segoe UI", 10, FontStyle.Bold), Color.Black, new Padding(0, 0, 0, 10));
    addText("\"Not used\" therefore means not used by the reports in that folder - not necessarily unused everywhere.",
        font, Color.Black, new Padding(0, 0, 0, 10));

    addText("Not detected - reports:", boldFont, Color.DarkRed, new Padding(0, 0, 0, 2));
    addText(
        "• Reports in the Power BI Service / Fabric that are not in your folder (other workspaces, apps, personal workspaces)\n" +
        "• Reports not saved in PBIR format (PBIX or PBIR-Legacy), paginated reports and dashboards\n" +
        "• DAX written in the report layer: report-level measures (reportExtensions.json) and visual calculations - e.g. a call to a DAX UDF\n" +
        "• Filters stored only in bookmarks (bookmarks/*.json)",
        font, Color.Black, new Padding(10, 0, 0, 10));

    addText("Not detected - other downstream use in the Service / Fabric:", boldFont, Color.DarkRed, new Padding(0, 0, 0, 2));
    addText(
        "• Copilot / AI: AI schema (Prep data for AI), verified answers, AI instructions, Fabric data agents, Q&A\n" +
        "• Translytical task flows / write-back, and Power Automate / Power Apps visuals\n" +
        "• Semantic Link / Semantic Link Labs notebooks, DAX queries via XMLA and external tools\n" +
        "• Excel (Analyze in Excel, PivotTables, connected tables) and composite models built on this model\n" +
        "• Subscriptions, data alerts and metrics / scorecards",
        font, Color.Black, new Padding(10, 0, 0, 10));

    addText("Before moving or deleting objects, check the model's lineage / impact analysis and usage metrics in the Service. " +
        "Deletions can be undone with Ctrl+Z in Tabular Editor until the model is saved.",
        font, Color.DimGray, new Padding(0, 0, 0, 12));

    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    var ok = new Button { Text = "I understand - continue", AutoSize = true, Font = boldFont, DialogResult = DialogResult.OK, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel, Padding = new Padding(6, 3, 6, 3) };
    buttons.Controls.AddRange(new Control[] { ok, cancel });
    layout.Controls.Add(buttons);

    form.Controls.Add(layout);
    form.AcceptButton = ok;
    form.CancelButton = cancel;
    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Dialog: select root folder and reports
// ---------------------------------------------------------------------------
string rootFolder = null;
using (var fbd = new FolderBrowserDialog())
{
    fbd.Description = "Select a folder containing one or more PBIR .Report folders (definition.pbir files).";
    if (fbd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no folder selected."); return; }
    rootFolder = fbd.SelectedPath;
}

var reportChoices = Directory.GetFiles(rootFolder, "definition.pbir", SearchOption.AllDirectories)
    .Select(f => new ReportChoice {
        DefinitionFile = f,
        ReportFolder = Path.GetDirectoryName(f),
        ReportName = CleanReportName(Path.GetDirectoryName(f)),
        ConnectionSummary = ReadConnectionSummary(f),
        Selected = true
    })
    .OrderBy(r => r.ReportName)
    .ThenBy(r => r.ReportFolder)
    .ToList();

if (reportChoices.Count == 0)
{
    Info("No definition.pbir files were found below:\n" + rootFolder);
    return;
}

using (var form = new Form())
{
    form.Text = "Select PBIR reports to scan";
    form.Size = BigDialogSize();
    form.MinimumSize = new Size(760, 520);
    form.StartPosition = FormStartPosition.CenterScreen;
    form.FormBorderStyle = FormBorderStyle.Sizable;
    form.MaximizeBox = true;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, Padding = new Padding(12) };
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var warning = new Label {
        Text = "This macro requires PBIR report files available on disk. Live-connected reports can be scanned for report references, but service-side model dependencies are taken from the model currently open in Tabular Editor.",
        Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(1100, 0), Font = font, ForeColor = Color.DarkSlateGray
    };
    var search = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 8, 0, 8) };
    var list = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = font, HorizontalScrollbar = true };
    var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var checkAll = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var clear = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { checkAll, clear, ok, cancel });

    bool suppress = false;
    Action rebuild = () =>
    {
        suppress = true;
        list.Items.Clear();
        var f = (search.Text ?? "").Trim();
        foreach (var r in reportChoices.Where(r => string.IsNullOrEmpty(f) || r.Label.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0))
            list.Items.Add(r.Label, r.Selected);
        suppress = false;
    };
    list.ItemCheck += (s, e) =>
    {
        if (suppress) return;
        var label = (string)list.Items[e.Index];
        var report = reportChoices.First(r => r.Label == label);
        report.Selected = e.NewValue == CheckState.Checked;
    };
    search.TextChanged += (s, e) => rebuild();
    checkAll.Click += (s, e) =>
    {
        foreach (var r in reportChoices.Where(r => string.IsNullOrEmpty(search.Text) || r.Label.IndexOf(search.Text, StringComparison.OrdinalIgnoreCase) >= 0))
            r.Selected = true;
        rebuild();
    };
    clear.Click += (s, e) =>
    {
        foreach (var r in reportChoices) r.Selected = false;
        rebuild();
    };
    ok.Click += (s, e) =>
    {
        if (!reportChoices.Any(r => r.Selected))
        {
            MessageBox.Show("Select at least one report.", "No reports selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    layout.Controls.Add(warning, 0, 0);
    layout.Controls.Add(search, 0, 1);
    layout.Controls.Add(list, 0, 2);
    layout.Controls.Add(buttons, 0, 3);
    form.Controls.Add(layout);
    form.AcceptButton = ok;
    form.CancelButton = cancel;
    rebuild();
    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

var selectedReports = reportChoices.Where(r => r.Selected).ToList();

// ---------------------------------------------------------------------------
// Dialog: choose object scope
// ---------------------------------------------------------------------------
bool includeMeasures = true;
bool includeColumns = true;
bool includeFunctions = true;
bool selectedOnly = false;

var selectedMeasures = Selected.Measures.ToList();
var selectedColumns = Selected.Columns.ToList();
var selectedFunctions = Selected.Functions.ToList();
bool hasSelectedObjects = selectedMeasures.Count > 0 || selectedColumns.Count > 0 || selectedFunctions.Count > 0;

using (var form = new Form())
{
    form.Text = "Choose model object scope";
    form.Width = 700;
    form.Height = 350;
    form.MinimumSize = new Size(620, 340);
    form.StartPosition = FormStartPosition.CenterScreen;
    form.FormBorderStyle = FormBorderStyle.Sizable;
    form.MaximizeBox = true;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 6, ColumnCount = 1, Padding = new Padding(12) };
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var info = new Label {
        Text = "Select what to classify. If you choose selected objects, only currently selected measures/columns in Tabular Editor are annotated. PBIR visual references are still scanned from the selected reports.",
        Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(640, 0), Font = font
    };
    var chkMeasures = new CheckBox { Text = "Classify measures", Checked = true, AutoSize = true, Font = font };
    var chkColumns = new CheckBox { Text = "Classify columns", Checked = true, AutoSize = true, Font = font };
    var chkFunctions = new CheckBox { Text = "Classify DAX UDFs", Checked = true, AutoSize = true, Font = font };
    var chkSelected = new CheckBox {
        Text = "Use selected objects only (" + selectedMeasures.Count + " measures, " + selectedColumns.Count + " columns, " + selectedFunctions.Count + " DAX UDFs selected)",
        Checked = hasSelectedObjects, Enabled = hasSelectedObjects, AutoSize = true, Font = font
    };
    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 12, 0, 0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.Add(ok);
    buttons.Controls.Add(cancel);
    ok.Click += (s, e) =>
    {
        if (!chkMeasures.Checked && !chkColumns.Checked && !chkFunctions.Checked)
        {
            MessageBox.Show("Select measures, columns, DAX UDFs, or a combination.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        includeMeasures = chkMeasures.Checked;
        includeColumns = chkColumns.Checked;
        includeFunctions = chkFunctions.Checked;
        selectedOnly = chkSelected.Checked;
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    layout.Controls.Add(info, 0, 0);
    layout.Controls.Add(chkMeasures, 0, 1);
    layout.Controls.Add(chkColumns, 0, 2);
    layout.Controls.Add(chkFunctions, 0, 3);
    layout.Controls.Add(chkSelected, 0, 4);
    layout.Controls.Add(buttons, 0, 5);
    form.Controls.Add(layout);
    form.AcceptButton = ok;
    form.CancelButton = cancel;
    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Scan PBIR references
// ---------------------------------------------------------------------------
var visualMeasureKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var visualColumnKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var visualReferenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
var scannedJsonFiles = 0;
var skippedJsonFiles = 0;

Action<string> addCount = key =>
{
    int count;
    visualReferenceCounts.TryGetValue(key, out count);
    visualReferenceCounts[key] = count + 1;
};

Action<JToken, string, Dictionary<string, string>, string> scanToken = null;
scanToken = (token, parentKey, aliasMap, reportName) =>
{
    var obj = token as JObject;
    if (obj != null)
    {
        var localAliases = new Dictionary<string, string>(aliasMap, StringComparer.OrdinalIgnoreCase);
        var fromArray = obj["From"] as JArray;
        if (fromArray != null)
        {
            foreach (var from in fromArray.OfType<JObject>())
            {
                var alias = (string)from["Name"];
                var entity = (string)from["Entity"];
                if (!string.IsNullOrEmpty(alias) && !string.IsNullOrEmpty(entity))
                    localAliases[alias] = entity;
            }
        }

        if (parentKey == "Measure" || parentKey == "Column")
        {
            var sourceRef = obj["Expression"]?["SourceRef"];
            var table = (string)sourceRef?["Entity"];
            if (table == null)
            {
                var source = (string)sourceRef?["Source"];
                if (source != null) localAliases.TryGetValue(source, out table);
            }
            var field = (string)obj["Property"];
            if (!string.IsNullOrEmpty(table) && !string.IsNullOrEmpty(field))
            {
                var key = table + "|" + field;
                if (parentKey == "Measure") visualMeasureKeys.Add(key); else visualColumnKeys.Add(key);
                addCount(parentKey + "|" + key);
            }
        }

        var selectorMetadata = (string)obj["selector"]?["metadata"];
        if (!string.IsNullOrWhiteSpace(selectorMetadata))
        {
            var dot = selectorMetadata.LastIndexOf('.');
            if (dot > 0 && dot < selectorMetadata.Length - 1)
            {
                var table = selectorMetadata.Substring(0, dot);
                var field = selectorMetadata.Substring(dot + 1);
                var key = table + "|" + field;
                visualMeasureKeys.Add(key);
                visualColumnKeys.Add(key);
                addCount("Selector|" + key);
            }
        }

        foreach (var property in obj.Properties())
            scanToken(property.Value, property.Name, localAliases, reportName);
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array) scanToken(item, parentKey, aliasMap, reportName);
    }
};

foreach (var report in selectedReports)
{
    var definitionFolder = Path.Combine(report.ReportFolder, "definition");
    if (!Directory.Exists(definitionFolder))
    {
        Warning("Skipping report without definition folder: " + report.ReportFolder);
        continue;
    }

    var jsonFiles = Directory.GetFiles(definitionFolder, "*.json", SearchOption.AllDirectories)
        .Where(f =>
            f.EndsWith("report.json", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith("page.json", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith("visual.json", StringComparison.OrdinalIgnoreCase))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    foreach (var file in jsonFiles)
    {
        try
        {
            var root = JObject.Parse(File.ReadAllText(file));
            scanToken(root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), report.ReportName);
            scannedJsonFiles++;
        }
        catch (Exception ex)
        {
            skippedJsonFiles++;
            Warning("Skipping unreadable PBIR JSON file:\n" + file + "\n\n" + ex.Message);
        }
    }
}

// ---------------------------------------------------------------------------
// Classify measures
// ---------------------------------------------------------------------------
var measureCandidates = selectedOnly
    ? selectedMeasures
    : Model.AllMeasures.Where(m => !IsExcludedSupportTable(m.Table)).ToList();
var columnCandidates = selectedOnly
    ? selectedColumns.Where(c => !IsExcludedSupportTable(c.Table)).ToList()
    : Model.AllColumns.Where(c => !IsExcludedSupportTable(c.Table)).ToList();
var functionCandidates = selectedOnly
    ? selectedFunctions
    : Model.Functions.ToList();
var measureClassificationCandidates = Model.AllMeasures.Where(m => !IsExcludedSupportTable(m.Table)).ToList();
var functionClassificationCandidates = Model.Functions.ToList();

var measureCategory = new Dictionary<Measure, string>();
var measureReasons = new Dictionary<Measure, string>();
var functionCategory = new Dictionary<Function, string>();
var functionReasons = new Dictionary<Function, string>();

foreach (var m in measureClassificationCandidates)
{
    var key = m.Table.Name + "|" + m.Name;
    if (visualMeasureKeys.Contains(key))
    {
        measureCategory[m] = CatVisual;
        measureReasons[m] = "Referenced directly by selected PBIR report visuals/filters.";
    }
}

bool changed = true;
while (changed)
{
    changed = false;
    foreach (var m in measureClassificationCandidates)
    {
        if (measureCategory.ContainsKey(m)) continue;
        var measureDeps = m.ReferencedBy.OfType<Measure>().ToList();
        var functionRefs = GetReferencedBy(m).OfType<Function>().ToList();
        if (measureDeps.Any(dep =>
            measureCategory.ContainsKey(dep) &&
            (measureCategory[dep] == CatVisual || measureCategory[dep] == CatMeasureWithVisual)) ||
            functionRefs.Any(dep =>
                functionCategory.ContainsKey(dep) &&
                functionCategory[dep] == CatDaxUdfWithVisual))
        {
            measureCategory[m] = CatMeasureWithVisual;
            measureReasons[m] = "A visually-used measure or DAX UDF depends on this measure.";
            changed = true;
        }
    }

    foreach (var f in functionClassificationCandidates)
    {
        if (functionCategory.ContainsKey(f)) continue;
        var referencedBy = GetReferencedBy(f).ToList();
        if (referencedBy.OfType<Measure>().Any(m =>
                measureCategory.ContainsKey(m) &&
                (measureCategory[m] == CatVisual || measureCategory[m] == CatMeasureWithVisual)) ||
            referencedBy.OfType<Function>().Any(fn =>
                functionCategory.ContainsKey(fn) &&
                functionCategory[fn] == CatDaxUdfWithVisual))
        {
            functionCategory[f] = CatDaxUdfWithVisual;
            functionReasons[f] = "A visually-used measure or DAX UDF depends on this DAX UDF.";
            changed = true;
        }
    }
}

foreach (var m in measureClassificationCandidates)
{
    if (measureCategory.ContainsKey(m)) continue;

    // For "is this measure used?", only downstream dependencies matter:
    // objects in ReferencedBy depend on this measure. DependsOn is what this
    // measure needs internally and should not keep an otherwise unused measure
    // out of NoDependencies.
    bool hasDaxUdfDependents = GetReferencedBy(m).OfType<Function>().Any();
    bool hasMeasureDependents = m.ReferencedBy.OfType<Measure>().Any();
    bool hasOtherModelDependents = m.ReferencedBy.Any(d => !(d is Measure));

    if (hasDaxUdfDependents)
    {
        measureCategory[m] = CatDaxUdfNoVisual;
        measureReasons[m] = "One or more DAX UDFs depend on this measure, but no selected PBIR visual dependency chain was found.";
    }
    else if (hasMeasureDependents || hasOtherModelDependents)
    {
        measureCategory[m] = CatMeasureNoVisual;
        measureReasons[m] = "One or more model objects depend on this measure, but no selected PBIR visual dependency chain was found.";
    }
    else
    {
        measureCategory[m] = CatNone;
        measureReasons[m] = "No selected PBIR visual dependency and no downstream model dependency was detected.";
    }
}

foreach (var f in functionClassificationCandidates)
{
    if (functionCategory.ContainsKey(f)) continue;
    // Same direction rule for DAX UDFs: a function is kept if another
    // function/measure depends on it. What the function depends on internally
    // does not make the function used.
    bool hasDaxUdfDependents = GetReferencedBy(f).OfType<Function>().Any();
    bool hasMeasureDependents = GetReferencedBy(f).OfType<Measure>().Any();
    if (hasDaxUdfDependents || hasMeasureDependents)
    {
        functionCategory[f] = CatDaxUdfNoVisual;
        functionReasons[f] = "One or more DAX UDFs/measures depend on this DAX UDF, but no selected PBIR visual dependency chain was found.";
    }
    else
    {
        functionCategory[f] = CatNone;
        functionReasons[f] = "No downstream measure dependency and no downstream DAX UDF dependency was detected.";
    }
}

// ---------------------------------------------------------------------------
// Classify columns
// ---------------------------------------------------------------------------
var columnCategory = new Dictionary<Column, string>();
var columnReasons = new Dictionary<Column, string>();

Func<Column, bool> HasRelationshipDependency = c =>
    Model.Relationships.Any(r => r.FromColumn == c || r.ToColumn == c);

Func<Column, bool> HasHierarchyDependency = c =>
    c.Table.Hierarchies.Any(h => h.Levels.Any(l => l.Column == c));

Func<Column, bool> HasSortByDependency = c =>
    c.Table.Columns.Any(other => other.SortByColumn == c || c.SortByColumn == other);

Func<Column, bool> HasCalendarOptionDependency = c =>
{
    try
    {
        if (!c.Table.Calendars.Any()) return false;
        foreach (var calendar in c.Table.Calendars)
        {
            foreach (var timeUnit in calendar.GetTimeUnits())
            {
                foreach (var prop in timeUnit.GetType().GetProperties())
                {
                    var value = prop.GetValue(timeUnit, null);
                    if (value == c) return true;

                    var enumerable = value as System.Collections.IEnumerable;
                    if (enumerable != null && !(value is string))
                        foreach (var item in enumerable)
                            if (item == c) return true;
                }
            }
        }
    }
    catch { }
    return false;
};

foreach (var c in columnCandidates)
{
    var key = c.Table.Name + "|" + c.Name;
    if (visualColumnKeys.Contains(key))
    {
        columnCategory[c] = CatVisual;
        columnReasons[c] = "Referenced directly by selected PBIR report visuals/filters.";
        continue;
    }

    var referencingMeasures = c.ReferencedBy.OfType<Measure>().ToList();
    var referencingFunctions = GetReferencedBy(c).OfType<Function>().ToList();
    bool usedByVisualMeasure = referencingMeasures.Any(m =>
        measureCategory.ContainsKey(m) &&
        (measureCategory[m] == CatVisual || measureCategory[m] == CatMeasureWithVisual));
    bool usedByDaxUdfMeasure = referencingMeasures.Any(m =>
        measureCategory.ContainsKey(m) &&
        measureCategory[m] == CatDaxUdfNoVisual) ||
        referencingFunctions.Any(f =>
            functionCategory.ContainsKey(f) &&
            (functionCategory[f] == CatDaxUdfWithVisual || functionCategory[f] == CatDaxUdfNoVisual));
    bool usedByNoVisualMeasure = referencingMeasures.Any(m =>
        measureCategory.ContainsKey(m) &&
        measureCategory[m] == CatMeasureNoVisual);

    if (usedByVisualMeasure)
    {
        columnCategory[c] = CatColumnUsedByVisualMeasure;
        columnReasons[c] = "Used by one or more measures with selected PBIR visual dependency.";
    }
    else if (usedByDaxUdfMeasure)
    {
        columnCategory[c] = CatColumnUsedByDaxUdfMeasure;
        columnReasons[c] = "Used by one or more measures with DAX UDF dependency but no selected PBIR visual dependency.";
    }
    else if (usedByNoVisualMeasure)
    {
        columnCategory[c] = CatColumnUsedByNoVisualMeasure;
        columnReasons[c] = "Used by measures that have model dependencies but no selected PBIR visual dependency.";
    }
    else if (HasRelationshipDependency(c) || HasHierarchyDependency(c) || HasSortByDependency(c) || HasCalendarOptionDependency(c) || c.ReferencedBy.Any())
    {
        columnCategory[c] = CatModelNoVisual;
        columnReasons[c] = "Has relationship, hierarchy, sort-by, Calendar option, or other model dependency without selected PBIR visual dependency.";
    }
    else
    {
        columnCategory[c] = CatNone;
        columnReasons[c] = "No selected PBIR visual dependency and no model dependency was detected.";
    }
}

// ---------------------------------------------------------------------------
// Preview results and choose action
// ---------------------------------------------------------------------------
var results = new List<ObjectResult>();
if (includeMeasures)
{
    foreach (var m in measureCandidates)
    {
        var key = "Measure|" + m.Table.Name + "|" + m.Name;
        int refs;
        visualReferenceCounts.TryGetValue(key, out refs);
        results.Add(new ObjectResult {
            Kind = "Measure",
            Table = m.Table.Name,
            Name = m.Name,
            Category = measureCategory[m],
            Reason = measureReasons[m],
            Hidden = m.IsHidden,
            VisualReferences = refs,
            DependsOnCount = m.DependsOn.Count(),
            ReferencedByCount = m.ReferencedBy.Count(),
            MeasureObject = m
        });
    }
}
if (includeColumns)
{
    foreach (var c in columnCandidates)
    {
        var key = "Column|" + c.Table.Name + "|" + c.Name;
        int refs;
        visualReferenceCounts.TryGetValue(key, out refs);
        results.Add(new ObjectResult {
            Kind = "Column",
            Table = c.Table.Name,
            Name = c.Name,
            Category = columnCategory[c],
            Reason = columnReasons[c],
            Hidden = c.IsHidden,
            VisualReferences = refs,
            DependsOnCount =
                (HasRelationshipDependency(c) ? 1 : 0) +
                (HasHierarchyDependency(c) ? 1 : 0) +
                (HasSortByDependency(c) ? 1 : 0) +
                (HasCalendarOptionDependency(c) ? 1 : 0),
            ReferencedByCount = c.ReferencedBy.Count(),
            ColumnObject = c
        });
    }
}
if (includeFunctions)
{
    foreach (var f in functionCandidates)
    {
        results.Add(new ObjectResult {
            Kind = "DAX UDF",
            Table = "",
            Name = f.Name,
            Category = functionCategory[f],
            Reason = functionReasons[f],
            Hidden = false,
            VisualReferences = 0,
            DependsOnCount = GetDependsOn(f).Count(),
            ReferencedByCount = GetReferencedBy(f).Count(),
            FunctionObject = f
        });
    }
}

// ---------------------------------------------------------------------------
// Duplicate / redundant detection (informational column in the output grid)
//   Measures: identical DAX (whitespace/case-insensitive) as another measure.
//   Columns:  same column name in other tables; same DAX expression (calculated)
//             or same source column (data) as another column in the same table.
// ---------------------------------------------------------------------------
Func<string, string> NormalizeDax = expr =>
    System.Text.RegularExpressions.Regex.Replace(expr ?? "", @"\s+", " ").Trim().ToUpperInvariant();
Func<IEnumerable<string>, string> ShortList = items =>
{
    var list = items.ToList();
    return string.Join(", ", list.Take(3)) + (list.Count > 3 ? " (+" + (list.Count - 3) + " more)" : "");
};

var duplicateText = new Dictionary<object, string>();

foreach (var g in Model.AllMeasures
    .Where(x => !string.IsNullOrWhiteSpace(x.Expression))
    .GroupBy(x => NormalizeDax(x.Expression))
    .Where(grp => grp.Count() > 1))
{
    foreach (var x in g)
        duplicateText[x] = "Same DAX as: " + ShortList(g.Where(o => o != x).Select(o => o.Table.Name + "[" + o.Name + "]"));
}

var columnNotes = new Dictionary<Column, List<string>>();
Action<Column, string> addColumnNote = (col, note) =>
{
    List<string> notes;
    if (!columnNotes.TryGetValue(col, out notes)) { notes = new List<string>(); columnNotes[col] = notes; }
    notes.Add(note);
};

foreach (var g in Model.AllColumns
    .Where(x => !IsExcludedSupportTable(x.Table))
    .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
    .Where(grp => grp.Select(x => x.Table).Distinct().Count() > 1))
{
    foreach (var x in g)
        addColumnNote(x, "Same name in: " + ShortList(g.Where(o => o.Table != x.Table).Select(o => o.Table.Name).Distinct()));
}

foreach (var g in Model.AllColumns
    .Where(x => !IsExcludedSupportTable(x.Table))
    .Select(x => new {
        Col = x,
        Key = x is CalculatedColumn
            ? "calc|" + NormalizeDax(((CalculatedColumn)x).Expression)
            : x is DataColumn && !string.IsNullOrWhiteSpace(((DataColumn)x).SourceColumn)
                ? "src|" + ((DataColumn)x).SourceColumn.ToUpperInvariant()
                : null
    })
    .Where(x => x.Key != null && x.Key != "calc|")
    .GroupBy(x => x.Col.Table.Name + "|" + x.Key)
    .Where(grp => grp.Count() > 1))
{
    foreach (var x in g)
        addColumnNote(x.Col, (x.Key.StartsWith("calc|") ? "Same DAX as: " : "Same source column as: ") +
            ShortList(g.Where(o => o.Col != x.Col).Select(o => "[" + o.Col.Name + "]")));
}
foreach (var kv in columnNotes) duplicateText[kv.Key] = string.Join("; ", kv.Value);

Func<ObjectResult, string> DuplicateOf = res =>
{
    string text;
    object key = (object)res.MeasureObject ?? res.ColumnObject;
    return key != null && duplicateText.TryGetValue(key, out text) ? text : "";
};

// ---------------------------------------------------------------------------
// Output grid (Tabular Editor's output window supports filtering/sorting)
// ---------------------------------------------------------------------------
var dt = new System.Data.DataTable();
dt.Columns.Add("Kind");
dt.Columns.Add("Table");
dt.Columns.Add("Name");
dt.Columns.Add("DependencyCategory");
dt.Columns.Add("Duplicate / redundant");
dt.Columns.Add("Display folder");
dt.Columns.Add("Hidden");
dt.Columns.Add("PBIR refs");
dt.Columns.Add("DependsOn (uses)");
dt.Columns.Add("ReferencedBy (used by)");
dt.Columns.Add("Reason");
foreach (var r in results.OrderBy(r => r.Kind).ThenBy(r => r.Category).ThenBy(r => r.Table).ThenBy(r => r.Name))
    dt.Rows.Add(
        r.Kind, r.Table, r.Name, r.Category, DuplicateOf(r),
        r.MeasureObject != null ? (r.MeasureObject.DisplayFolder ?? "") : r.ColumnObject != null ? (r.ColumnObject.DisplayFolder ?? "") : "",
        r.Hidden ? "Yes" : "No", r.VisualReferences, r.DependsOnCount, r.ReferencedByCount, r.Reason);
Output(dt);

// ---------------------------------------------------------------------------
// Choose action (annotations are applied by every option except Do nothing)
// ---------------------------------------------------------------------------
Func<string, bool> IsUnusedCategory = cat => cat == CatNone || cat == CatMeasureNoVisual || cat == CatDaxUdfNoVisual;
Func<string, bool> IsLegacyUnusedFolder = folder =>
    folder == LegacyUnusedFolder || folder.EndsWith("\\" + LegacyUnusedFolder);
Func<Measure, bool> IsInUnusedFolder = x =>
{
    var folder = x.DisplayFolder ?? "";
    return folder == UnusedFolder || IsLegacyUnusedFolder(folder);
};
// Folder to return to: stored annotation for the new root folder, or the
// parent path for folders created by older versions of this macro.
Func<Measure, string> OriginalFolderOf = x =>
{
    var folder = x.DisplayFolder ?? "";
    if (folder == UnusedFolder) return x.GetAnnotation(OriginalFolderAnnotation) ?? "";
    if (folder == LegacyUnusedFolder) return "";
    if (folder.EndsWith("\\" + LegacyUnusedFolder)) return folder.Substring(0, folder.Length - LegacyUnusedFolder.Length - 1);
    return folder;
};

var actionCandidates = results
    .Where(r => r.Kind == "Measure" && r.MeasureObject != null && IsUnusedCategory(r.Category))
    .OrderBy(r => r.Table).ThenBy(r => r.Name)
    .ToList();

// Measures in the "not to be used" folder that are now used again (e.g. a
// visual was created) are moved back out of that folder.
var measuresToRestore = results
    .Where(r => r.Kind == "Measure" && r.MeasureObject != null && !IsUnusedCategory(r.Category) && IsInUnusedFolder(r.MeasureObject))
    .Select(r => r.MeasureObject)
    .ToList();

var summary = string.Join("\n",
    results.GroupBy(r => r.Kind + " - " + r.Category)
        .OrderBy(grp => grp.Key)
        .Select(grp => grp.Key + ": " + grp.Count()));
int duplicateCount = results.Count(r => DuplicateOf(r) != "");

string chosenAction = "Nothing";
using (var form = new Form())
{
    form.Text = "Apply dependency classification";
    form.AutoSize = true;
    form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.MaximizeBox = false;
    form.MinimizeBox = false;
    form.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 4, Padding = new Padding(12) };
    var heading = new Label {
        Text = "What should be done with the classification?",
        AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
    };
    var details = new Label {
        Text =
            "Reports included: " + selectedReports.Count + "  |  PBIR JSON files scanned: " + scannedJsonFiles +
            (skippedJsonFiles > 0 ? " (" + skippedJsonFiles + " skipped)" : "") + "\n\n" +
            summary + "\n\n" +
            "Duplicate / redundant objects flagged: " + duplicateCount + " (see the output grid)\n" +
            "Not-used measures (" + CatNone + ", " + CatMeasureNoVisual + ", " + CatDaxUdfNoVisual + "): " + actionCandidates.Count +
            (measuresToRestore.Count > 0
                ? "\nMeasures in '" + UnusedFolder + "' that are used again and will be moved back to their original folder: " + measuresToRestore.Count
                : ""),
        AutoSize = true, MaximumSize = new Size(760, 0), Font = font, Margin = new Padding(0, 0, 0, 10)
    };
    var note = new Label {
        Text = "Every option except Do nothing writes the " + AnnotationName + " annotations" +
            (measuresToRestore.Count > 0 ? " and moves re-used measures back out of '" + UnusedFolder + "'" : "") + ".",
        AutoSize = true, MaximumSize = new Size(760, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 12)
    };
    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    Func<string, string, bool, Button> makeButton = (text, action, enabled) =>
    {
        var button = new Button { Text = text, AutoSize = true, Font = font, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0), Enabled = enabled };
        button.Click += (s, e) => { chosenAction = action; form.DialogResult = DialogResult.OK; form.Close(); };
        buttonRow.Controls.Add(button);
        return button;
    };
    bool hasCandidates = includeMeasures && actionCandidates.Count > 0;
    makeButton("Move to '" + UnusedFolder + "'", "Move", hasCandidates);
    makeButton("Choose measures to delete...", "Delete", hasCandidates);
    var annotateButton = makeButton("Keep annotations only", "Annotate", true);
    var nothingButton = new Button { Text = "Do nothing", AutoSize = true, Font = font, Padding = new Padding(6, 3, 6, 3), DialogResult = DialogResult.Cancel };
    buttonRow.Controls.Add(nothingButton);

    layout.Controls.Add(heading, 0, 0);
    layout.Controls.Add(details, 0, 1);
    layout.Controls.Add(note, 0, 2);
    layout.Controls.Add(buttonRow, 0, 3);
    form.Controls.Add(layout);
    form.CancelButton = nothingButton;
    form.Shown += (s, e) => annotateButton.Focus();
    form.ShowDialog();
}

if (chosenAction == "Nothing") { Info("Do nothing - no annotations or model changes were applied."); return; }

foreach (var r in results)
{
    if (r.MeasureObject != null) r.MeasureObject.SetAnnotation(AnnotationName, r.Category);
    if (r.ColumnObject != null) r.ColumnObject.SetAnnotation(AnnotationName, r.Category);
    if (r.FunctionObject != null) r.FunctionObject.SetAnnotation(AnnotationName, r.Category);
}

foreach (var x in measuresToRestore)
{
    x.DisplayFolder = OriginalFolderOf(x);
    x.RemoveAnnotation(OriginalFolderAnnotation);
}
var restoredText = measuresToRestore.Count > 0
    ? "\nMoved " + measuresToRestore.Count + " re-used measure(s) back out of '" + UnusedFolder + "'."
    : "";

// ---------------------------------------------------------------------------
// Move: one root display folder per table, original folder kept as annotation
// ---------------------------------------------------------------------------
if (chosenAction == "Move")
{
    int moved = 0;
    foreach (var r in actionCandidates)
    {
        var x = r.MeasureObject;
        if ((x.DisplayFolder ?? "") == UnusedFolder) continue;
        var original = OriginalFolderOf(x);
        if (string.IsNullOrEmpty(original)) x.RemoveAnnotation(OriginalFolderAnnotation);
        else x.SetAnnotation(OriginalFolderAnnotation, original);
        x.DisplayFolder = UnusedFolder;
        moved++;
    }
    Info("Done. Applied annotations and moved " + moved + " measure(s) to the root display folder '" + UnusedFolder + "'" +
        (actionCandidates.Count > moved ? " (" + (actionCandidates.Count - moved) + " were already there)" : "") + "." + restoredText +
        "\n\nRun the macro again later: measures that are used by a visual by then are moved back to their original folder.");
    return;
}

// ---------------------------------------------------------------------------
// Delete: pick measures from a fully expanded tree
// ---------------------------------------------------------------------------
if (chosenAction == "Delete")
{
    var deleteSet = new HashSet<Measure>();
    using (var form = new Form())
    {
        form.Text = "Select measures to delete";
        form.Size = BigDialogSize();
        form.MinimumSize = new Size(760, 520);
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.Sizable;
        form.MaximizeBox = true;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var label = new Label {
            Text = "Select measures to delete. Unchecked measures keep their DependencyCategory annotation only.",
            Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        var filterPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true, Margin = new Padding(0, 8, 0, 8) };
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var searchLabel = new Label { Text = "Search:", AutoSize = true, Font = font, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) };
        var searchBox = new TextBox { Dock = DockStyle.Fill, Font = font };
        var categoryValues = new[] { CatNone, CatDaxUdfNoVisual, CatMeasureNoVisual };
        var categoryList = new CheckedListBox { CheckOnClick = true, Font = font, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle };
        foreach (var cat in categoryValues) categoryList.Items.Add(cat, true);
        // Size the category list to its content so no scrollbar is needed
        categoryList.Height = categoryList.ItemHeight * categoryValues.Length + 6;
        categoryList.Width = categoryValues.Max(v => TextRenderer.MeasureText(v, font).Width) + 40;
        filterPanel.Controls.Add(searchLabel, 0, 0);
        filterPanel.Controls.Add(searchBox, 1, 0);
        filterPanel.Controls.Add(new Label { Text = "Categories:", AutoSize = true, Font = font, Anchor = AnchorStyles.Left, Margin = new Padding(8, 6, 8, 0) }, 2, 0);
        filterPanel.Controls.Add(categoryList, 3, 0);

        var tree = new TreeView { Dock = DockStyle.Fill, CheckBoxes = true, Font = font };
        var lblStatus = new Label { Text = "0 measure(s) selected for deletion", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisible = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearAll = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var ok = new Button { Text = "Delete selected", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { checkVisible, clearAll, ok, cancel });

        Action updateStatus = () => lblStatus.Text = deleteSet.Count + " measure(s) selected for deletion";
        bool suppress = false;
        Action updateTree = () =>
        {
            suppress = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();
            var search = (searchBox.Text ?? "").Trim();
            var allowedCategories = new HashSet<string>(categoryList.CheckedItems.Cast<string>(), StringComparer.OrdinalIgnoreCase);
            var filtered = actionCandidates.Where(r =>
                allowedCategories.Contains(r.Category) &&
                (string.IsNullOrEmpty(search) ||
                 r.Table.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 r.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (r.MeasureObject.DisplayFolder ?? "").IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0))
                .ToList();
            foreach (var tableGroup in filtered.GroupBy(r => r.Table).OrderBy(grp => grp.Key))
            {
                var tableNode = new TreeNode(tableGroup.Key);
                tree.Nodes.Add(tableNode);
                foreach (var folderGroup in tableGroup.GroupBy(r => string.IsNullOrWhiteSpace(r.MeasureObject.DisplayFolder) ? "(no display folder)" : r.MeasureObject.DisplayFolder).OrderBy(grp => grp.Key))
                {
                    var folderNode = new TreeNode(folderGroup.Key);
                    tableNode.Nodes.Add(folderNode);
                    foreach (var r in folderGroup.OrderBy(o => o.Name))
                    {
                        var dup = DuplicateOf(r);
                        folderNode.Nodes.Add(new TreeNode(r.Name + "  [" + r.Category + "]" + (dup != "" ? "  - " + dup : "")) {
                            Tag = r.MeasureObject,
                            Checked = deleteSet.Contains(r.MeasureObject)
                        });
                    }
                }
            }
            tree.ExpandAll();
            if (tree.Nodes.Count > 0) tree.Nodes[0].EnsureVisible();
            tree.EndUpdate();
            suppress = false;
            updateStatus();
        };

        Action<TreeNode, bool> setCheckedRecursive = null;
        setCheckedRecursive = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var measure = node.Tag as Measure;
            if (measure != null)
            {
                if (isChecked) deleteSet.Add(measure); else deleteSet.Remove(measure);
            }
            foreach (TreeNode child in node.Nodes) setCheckedRecursive(child, isChecked);
        };
        tree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;
            setCheckedRecursive(e.Node, e.Node.Checked);
            suppress = false;
            updateStatus();
        };
        searchBox.TextChanged += (s, e) => updateTree();
        categoryList.ItemCheck += (s, e) => form.BeginInvoke((Action)(() => updateTree()));
        checkVisible.Click += (s, e) =>
        {
            suppress = true;
            foreach (TreeNode node in tree.Nodes) setCheckedRecursive(node, true);
            suppress = false;
            updateTree();
        };
        clearAll.Click += (s, e) =>
        {
            deleteSet.Clear();
            updateTree();
        };
        ok.Click += (s, e) =>
        {
            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        layout.Controls.Add(label, 0, 0);
        layout.Controls.Add(filterPanel, 0, 1);
        layout.Controls.Add(tree, 0, 2);
        layout.Controls.Add(lblStatus, 0, 3);
        layout.Controls.Add(buttons, 0, 4);
        form.Controls.Add(layout);
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        updateTree();
        if (form.ShowDialog() == DialogResult.OK && deleteSet.Count > 0)
        {
            var confirmDelete = MessageBox.Show(
                "Delete " + deleteSet.Count + " selected measure(s)?\n\nThis changes the semantic model (undo with Ctrl+Z in Tabular Editor). Test on a copy/dev model first.",
                "Confirm measure deletion",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirmDelete == DialogResult.Yes)
            {
                int deleted = deleteSet.Count;
                foreach (var x in deleteSet.ToList()) x.Delete();
                Info("Done. Applied annotations and deleted " + deleted + " measure(s)." + restoredText);
                return;
            }
        }
    }
    Info("No measures deleted. Applied DependencyCategory annotations to " + results.Count + " object(s)." + restoredText);
    return;
}

Info("Done. Applied DependencyCategory annotations to " + results.Count + " object(s)." + restoredText);
