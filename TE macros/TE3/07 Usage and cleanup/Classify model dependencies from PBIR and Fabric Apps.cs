// ============================================================================
// 22 - Classify model dependencies from PBIR and Fabric Apps.cs
//
// READ-ONLY overview of how the open model is used by PBIR reports and/or Fabric
// Apps (Rayfin projects). It ends with the output grid - nothing in the model is
// changed. Organizing the result is left to dedicated macros (19 for measures).
// (Replaces macro 16, which was the PBIR-only version with move/delete actions.)
//
// Fabric App queries count as direct use exactly like PBIR visuals:
// a measure or column that an app queries is a VisualDependency, and the Reason
// column says whether the use comes from reports, apps or both.
//
// Sources and folders:
//   - First dialog: scan PBIR reports, Fabric Apps or both, each with its folder.
//     PBIR can be limited to reports whose definition.pbir points to the open
//     model; apps can be limited by a model filter (connector item id,
//     workspace id or alias from rayfin/rayfin.yml connectors or the app's
//     fabric.yaml semanticModels). Without a filter every app is scanned.
//   - When the model was opened from disk, the folders default to the
//     repository root above the model, so no browsing is needed.
//   - Settings are remembered per model in
//     %APPDATA%\TabularEditor3\MacroSettings\dependency-scan-sources.json.
//     With "start straight away" ticked, the next run skips the source and
//     selection dialogs; hold Shift when starting the macro to show them again.
//   - Fabric App scanning: .dax files are parsed for Table[Column],
//     'Table'[Column], Table[Measure] and [Measure]; .ts/.tsx files only add
//     names that exist in the model. Query result aliases ("Week" / [Week]) and
//     query measures (DEFINE MEASURE) are not model objects and are ignored;
//     other unknown names are listed as app references not found in the model.
//
// Measures, columns, DAX UDFs, field parameters and calculation groups of the
// open model are classified (DependencyCategory column in the output grid) as:
//   DependencyCategory = VisualDependency
//   DependencyCategory = MeasureWithVisualDependency
//   DependencyCategory = DaxUdfWithVisualDependency
//   DependencyCategory = DaxUdfDependencyWithoutVisualDependency
//   DependencyCategory = MeasureDependencyWithoutVisualDependency
//   DependencyCategory = ColumnUsedByDaxUdfDependency
//   DependencyCategory = ColumnUsedByVisualDependentMeasure
//   DependencyCategory = ColumnUsedByMeasureDependencyWithoutVisualDependency
//   DependencyCategory = ModelDependencyWithoutVisualDependency
//   DependencyCategory = PerspectiveDependencyWithoutVisualDependency
//   DependencyCategory = NoDependencies
//
// Field parameters, calculation groups and perspectives:
//   - A measure or column listed in a field parameter that is used in the
//     reports is a VisualDependency (it is shown in visuals through it).
//   - A measure, column or DAX UDF referenced by a calculation item of a
//     calculation group used in the reports counts as used by a visually-used
//     measure. A DAX UDF referenced only by calculation items of calculation
//     groups NOT used in the reports (or by calculated columns / tables, RLS)
//     is ModelDependencyWithoutVisualDependency.
//   - Field parameters and calculation groups get their own rows:
//     VisualDependency when used in the reports.
//   - A measure in a perspective, or used by a measure in a perspective, is
//     PerspectiveDependencyWithoutVisualDependency: a model dependency for
//     downstream use (e.g. perspective models), not "not used". Columns in a perspective (or used by such
//     measures) are a model dependency too. The Perspectives column in the
//     output lists the membership for every object.
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
//   7. A summary message. That is the last step: no annotations, folders or
//      deletions. Measures still in the 'XX - Not to be used' folder from older
//      versions of this macro / macro 16 are cleaned up by macro 19.
//
// Notes / limits for this first version:
//   - Requires PBIR files to be available on disk.
//   - The open Tabular Editor model is the source of truth for model
//     dependencies. Live-connected report metadata is listed when found in
//     definition.pbir, but service-side model dependencies are not fetched.
//   - Read-only: neither the model nor the PBIR / app files are changed.
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

const string CatVisual = "VisualDependency";
const string CatMeasureWithVisual = "MeasureWithVisualDependency";
const string CatDaxUdfWithVisual = "DaxUdfWithVisualDependency";
const string CatDaxUdfNoVisual = "DaxUdfDependencyWithoutVisualDependency";
const string CatMeasureNoVisual = "MeasureDependencyWithoutVisualDependency";
const string CatColumnUsedByVisualMeasure = "ColumnUsedByVisualDependentMeasure";
const string CatColumnUsedByDaxUdfMeasure = "ColumnUsedByDaxUdfDependency";
const string CatColumnUsedByNoVisualMeasure = "ColumnUsedByMeasureDependencyWithoutVisualDependency";
const string CatModelNoVisual = "ModelDependencyWithoutVisualDependency";
const string CatPerspective = "PerspectiveDependencyWithoutVisualDependency";
const string CatNone = "NoDependencies";

class ReportChoice
{
    public string Kind = "PBIR";   // PBIR or App
    public string DefinitionFile;
    public string ReportFolder;
    public string ReportName;
    public string ConnectionSummary;
    public bool Selected;
    public string Label
    {
        get { return "[" + Kind + "] " + ReportName + "  |  " + ConnectionSummary + "  |  " + ReportFolder; }
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
    public Table TableObject;
    public string Perspectives;
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

    addText("This analysis only sees the model open in Tabular Editor and the PBIR reports and Fabric Apps in the folders you select.",
        new Font("Segoe UI", 10, FontStyle.Bold), Color.Black, new Padding(0, 0, 0, 10));
    addText("\"Not used\" therefore means not used by those reports and apps - not necessarily unused everywhere.",
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

    addText("This macro is read-only: it ends with the output grid. Before acting on the result (organizing, hiding or deleting), " +
        "check the model's lineage / impact analysis and usage metrics in the Service.",
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
// Sources: PBIR reports and/or Fabric Apps. Folders and options are remembered
// per model, so later runs can start without any folder dialog.
// ---------------------------------------------------------------------------
string[] SkippedFolders = { "node_modules", "dist", ".agents", ".git", ".temp", ".vite", "coverage", "bin", "obj" };

// Walks a folder tree without descending into skipped folders (node_modules alone
// can hold tens of thousands of files) or folders that cannot be read.
Func<string, IEnumerable<string>> EnumerateSourceFiles = root =>
{
    var result = new List<string>();
    var pending = new Stack<string>();
    pending.Push(root);
    while (pending.Count > 0)
    {
        var dir = pending.Pop();
        try
        {
            result.AddRange(Directory.GetFiles(dir));
            foreach (var sub in Directory.GetDirectories(dir))
                if (!SkippedFolders.Contains(Path.GetFileName(sub), StringComparer.OrdinalIgnoreCase))
                    pending.Push(sub);
        }
        catch (UnauthorizedAccessException) { }
        catch (PathTooLongException) { }
    }
    return result;
};

// Where the open model was loaded from (a .bim file or a TMDL / PBIP folder), when
// Tabular Editor exposes it. Empty when connected to a server.
string modelSourcePath = null;
try
{
    var handler = Model.GetType().GetProperty("Handler")?.GetValue(Model, null);
    modelSourcePath = handler?.GetType().GetProperty("Source")?.GetValue(handler, null) as string;
    if (modelSourcePath != null && !File.Exists(modelSourcePath) && !Directory.Exists(modelSourcePath)) modelSourcePath = null;
}
catch { }

// The .SemanticModel folder of the open model, and the repository root above it.
string modelFolder = null, repoRoot = null;
if (modelSourcePath != null)
{
    var startDir = File.Exists(modelSourcePath) ? Path.GetDirectoryName(modelSourcePath) : modelSourcePath;
    for (var d = startDir; d != null; d = Path.GetDirectoryName(d))
    {
        if (modelFolder == null && d.EndsWith(".SemanticModel", StringComparison.OrdinalIgnoreCase)) modelFolder = d;
        if (Directory.Exists(Path.Combine(d, ".git"))) { repoRoot = d; break; }
    }
    if (repoRoot == null) repoRoot = modelFolder != null ? Path.GetDirectoryName(modelFolder) : startDir;
}

string modelName = "";
try { modelName = Model.Database?.Name ?? ""; } catch { }
string modelKey = modelSourcePath ?? modelName;
try
{
    // Database.ID through reflection: not every Tabular Editor version exposes it.
    var db = Model.Database;
    var dbId = db == null ? null : db.GetType().GetProperty("ID")?.GetValue(db, null) as string;
    if (modelSourcePath == null && !string.IsNullOrEmpty(dbId)) modelKey = modelName + "|" + dbId;
}
catch { }

var settingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
    "TabularEditor3", "MacroSettings", "dependency-scan-sources.json");
JObject allSettings;
try { allSettings = File.Exists(settingsFile) ? JObject.Parse(File.ReadAllText(settingsFile)) : new JObject(); }
catch { allSettings = new JObject(); }
var saved = allSettings[modelKey] as JObject ?? new JObject();

bool scanPbir = (bool?)saved["scanPbir"] ?? true;
bool scanApps = (bool?)saved["scanApps"] ?? false;
string pbirFolder = (string)saved["pbirFolder"] ?? repoRoot ?? "";
string appsFolder = (string)saved["appsFolder"] ?? repoRoot ?? "";
bool onlyThisModelReports = ((bool?)saved["onlyThisModelReports"] ?? true) && modelFolder != null;
string appModelFilter = (string)saved["appModelFilter"] ?? "";
bool rememberSources = true;
bool skipSourceDialog = (bool?)saved["skipSourceDialog"] ?? false;

bool autoRun = skipSourceDialog
    && (Control.ModifierKeys & Keys.Shift) == 0
    && (!scanPbir || Directory.Exists(pbirFolder))
    && (!scanApps || Directory.Exists(appsFolder))
    && (scanPbir || scanApps);

if (!autoRun)
{
    using (var form = new Form())
    {
        form.Text = "Choose what to scan";
        form.AutoSize = true;
        form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MaximizeBox = false;
        form.MinimizeBox = false;
        form.StartPosition = FormStartPosition.CenterScreen;
        var font = new Font("Segoe UI", 9);
        var boldFont = new Font("Segoe UI", 9, FontStyle.Bold);
        form.Font = font;

        var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3, Padding = new Padding(14) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 480));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        Action<Control, int> addSpan = (c, span) => { layout.Controls.Add(c); layout.SetColumnSpan(c, span); };

        addSpan(new Label {
            Text = "Model: " + (modelName == "" ? "(unnamed)" : modelName) + "\n" +
                   (modelSourcePath != null ? "Loaded from: " + modelSourcePath : "Not loaded from a file - pick folders below; they are remembered for this model."),
            AutoSize = true, MaximumSize = new Size(720, 0), Margin = new Padding(0, 0, 0, 10)
        }, 3);

        Func<string, string, TextBox> folderRow = (label, value) =>
        {
            layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(18, 8, 8, 0) });
            var box = new TextBox { Text = value, Width = 470, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            layout.Controls.Add(box);
            var browse = new Button { Text = "Browse…", AutoSize = true };
            browse.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog { SelectedPath = Directory.Exists(box.Text) ? box.Text : "" })
                    if (fbd.ShowDialog() == DialogResult.OK) box.Text = fbd.SelectedPath;
            };
            layout.Controls.Add(browse);
            return box;
        };

        var chkPbir = new CheckBox { Text = "PBIR reports (definition.pbir)", Checked = scanPbir, AutoSize = true, Font = boldFont, Margin = new Padding(0, 6, 0, 0) };
        addSpan(chkPbir, 3);
        var pbirBox = folderRow("Folder", pbirFolder);
        var chkOnlyThisModel = new CheckBox {
            Text = modelFolder != null
                ? "Only reports whose definition.pbir points to this model (" + Path.GetFileName(modelFolder) + ")"
                : "Only reports pointing to this model (needs a model opened from a .SemanticModel folder)",
            Checked = onlyThisModelReports, AutoSize = true, Margin = new Padding(18, 4, 0, 6)
        };
        addSpan(chkOnlyThisModel, 3);

        var chkApps = new CheckBox { Text = "Fabric Apps (rayfin/rayfin.yml)", Checked = scanApps, AutoSize = true, Font = boldFont, Margin = new Padding(0, 10, 0, 0) };
        addSpan(chkApps, 3);
        var appsBox = folderRow("Folder", appsFolder);
        layout.Controls.Add(new Label { Text = "Model filter", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(18, 8, 8, 0) });
        var filterBox = new TextBox { Text = appModelFilter, Width = 470, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        layout.Controls.Add(filterBox);
        layout.Controls.Add(new Label { Text = "optional", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 0) });
        addSpan(new Label {
            Text = "Semantic model item id, workspace id or alias from rayfin.yml / fabric.yaml. Empty = every app below the folder.",
            AutoSize = true, MaximumSize = new Size(700, 0), ForeColor = Color.DimGray, Margin = new Padding(18, 2, 0, 10)
        }, 3);

        var chkRemember = new CheckBox { Text = "Remember these settings for this model", Checked = true, AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
        addSpan(chkRemember, 3);
        var chkSkip = new CheckBox {
            Text = "Next time, start straight away with these settings (hold Shift when starting the macro to see this dialog again)",
            Checked = skipSourceDialog, AutoSize = true, MaximumSize = new Size(720, 0), Margin = new Padding(0, 2, 0, 10)
        };
        addSpan(chkSkip, 3);
        chkRemember.CheckedChanged += (s, e) => { chkSkip.Enabled = chkRemember.Checked; if (!chkRemember.Checked) chkSkip.Checked = false; };

        Action syncEnabled = () =>
        {
            pbirBox.Enabled = chkPbir.Checked;
            chkOnlyThisModel.Enabled = chkPbir.Checked && modelFolder != null;
            appsBox.Enabled = chkApps.Checked;
            filterBox.Enabled = chkApps.Checked;
        };
        chkPbir.CheckedChanged += (s, e) => syncEnabled();
        chkApps.CheckedChanged += (s, e) => syncEnabled();
        syncEnabled();

        var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
        var ok = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { ok, cancel });
        addSpan(buttons, 3);
        ok.Click += (s, e) =>
        {
            string problem = null;
            if (!chkPbir.Checked && !chkApps.Checked) problem = "Select PBIR reports, Fabric Apps, or both.";
            else if (chkPbir.Checked && !Directory.Exists(pbirBox.Text.Trim().Trim('"'))) problem = "The PBIR folder does not exist.";
            else if (chkApps.Checked && !Directory.Exists(appsBox.Text.Trim().Trim('"'))) problem = "The Fabric Apps folder does not exist.";
            if (problem != null) { MessageBox.Show(problem, "Check the inputs", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        form.Controls.Add(layout);
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }

        scanPbir = chkPbir.Checked;
        scanApps = chkApps.Checked;
        pbirFolder = pbirBox.Text.Trim().Trim('"');
        appsFolder = appsBox.Text.Trim().Trim('"');
        onlyThisModelReports = chkOnlyThisModel.Checked && modelFolder != null;
        appModelFilter = filterBox.Text.Trim();
        rememberSources = chkRemember.Checked;
        skipSourceDialog = chkSkip.Checked;
    }
}

if (rememberSources)
{
    var entry = new JObject();
    entry["scanPbir"] = scanPbir;
    entry["pbirFolder"] = pbirFolder;
    entry["onlyThisModelReports"] = onlyThisModelReports;
    entry["scanApps"] = scanApps;
    entry["appsFolder"] = appsFolder;
    entry["appModelFilter"] = appModelFilter;
    entry["skipSourceDialog"] = skipSourceDialog;
    allSettings[modelKey] = entry;
    try
    {
        Directory.CreateDirectory(Path.GetDirectoryName(settingsFile));
        File.WriteAllText(settingsFile, allSettings.ToString());
    }
    catch (Exception ex) { Warning("Could not save the scan settings:\n" + ex.Message); }
}

// Does a definition.pbir point to the open model? (byPath is relative to the report folder)
Func<string, bool> PointsToOpenModel = definitionFile =>
{
    if (modelFolder == null) return true;
    try
    {
        var root = JObject.Parse(File.ReadAllText(definitionFile));
        var dataset = root["datasetReference"] ?? root["semanticModelReference"];
        var byPath = (string)dataset?["byPath"]?["path"];
        if (string.IsNullOrWhiteSpace(byPath)) return false;
        var target = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(definitionFile), byPath)).TrimEnd('\\', '/');
        return string.Equals(target, modelFolder.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }
    catch { return false; }
};

// Semantic-model connectors from the top-level "connectors:" block of rayfin.yml
// (services.connectors is only the on/off switch). Each entry: alias, workspaceId, itemId.
Func<string, List<string[]>> ReadAppConnectors = rayfinYml =>
{
    var list = new List<string[]>();
    bool inBlock = false;
    string alias = null, type = null, ws = null, item = null;
    Action flush = () =>
    {
        if (alias != null && string.Equals(type, "fabric-semanticmodel", StringComparison.OrdinalIgnoreCase))
            list.Add(new[] { alias, ws ?? "", item ?? "" });
        alias = null; type = null; ws = null; item = null;
    };
    foreach (var raw in File.ReadAllLines(rayfinYml))
    {
        if (raw.Length > 0 && !char.IsWhiteSpace(raw[0]))
        {
            if (inBlock) flush();
            inBlock = raw.TrimEnd() == "connectors:";
            continue;
        }
        if (!inBlock) continue;
        var line = raw.Trim();
        Func<int, string> value = n => line.Substring(n).Trim().Trim('"', '\'');
        if (line.StartsWith("- name:")) { flush(); alias = value(7); }
        else if (line.StartsWith("type:")) type = value(5);
        else if (line.StartsWith("workspaceId:")) ws = value(12);
        else if (line.StartsWith("itemId:")) item = value(7);
    }
    if (inBlock) flush();
    return list;
};

// Semantic models declared in the app's fabric.yaml (fabric-app-data, used by the
// Data App template) for the active profile: profiles.<p>.semanticModels.<alias>.
Func<string, List<string[]>> ReadFabricYaml = appRoot =>
{
    var list = new List<string[]>();
    var file = new[] { "fabric.yaml", "fabric.yml" }.Select(n => Path.Combine(appRoot, n)).FirstOrDefault(File.Exists);
    if (file == null) return list;
    var lines = File.ReadAllLines(file);
    string active = lines.Where(l => l.StartsWith("activeProfile:"))
        .Select(l => l.Substring(14).Trim().Trim('"', '\'')).FirstOrDefault();
    bool inProfiles = false, inModels = false;
    int profileIndent = -1, modelsIndent = -1, aliasIndent = -1;
    string profile = null, alias = null, ws = null, item = null;
    Action flush = () =>
    {
        if (alias != null && (string.IsNullOrEmpty(active) || profile == active)) list.Add(new[] { alias, ws ?? "", item ?? "" });
        alias = null; ws = null; item = null;
    };
    foreach (var raw in lines)
    {
        var text = raw.Trim();
        if (text.Length == 0 || text.StartsWith("#")) continue;
        int indent = raw.Length - raw.TrimStart().Length;
        int colon = text.IndexOf(':');
        var key = (colon < 0 ? text : text.Substring(0, colon)).Trim().Trim('"', '\'');
        var val = colon < 0 ? "" : text.Substring(colon + 1).Trim().Trim('"', '\'');
        if (inModels && indent <= modelsIndent) { flush(); inModels = false; }
        if (indent == 0) { inProfiles = key == "profiles"; profile = null; profileIndent = -1; continue; }
        if (!inProfiles) continue;
        if (profileIndent < 0 || indent <= profileIndent) { profileIndent = indent; profile = key; continue; }
        if (!inModels)
        {
            if (key == "semanticModels") { inModels = true; modelsIndent = indent; aliasIndent = -1; }
            continue;
        }
        if (aliasIndent < 0 || indent <= aliasIndent) { flush(); aliasIndent = indent; alias = key; continue; }
        if (key == "workspaceId") ws = val;
        else if (key == "itemId") item = val;
    }
    flush();
    return list;
};

var reportChoices = new List<ReportChoice>();
int reportsNotForThisModel = 0, appsNotMatchingFilter = 0;

if (scanPbir)
{
    foreach (var f in EnumerateSourceFiles(pbirFolder).Where(p => string.Equals(Path.GetFileName(p), "definition.pbir", StringComparison.OrdinalIgnoreCase)))
    {
        if (onlyThisModelReports && !PointsToOpenModel(f)) { reportsNotForThisModel++; continue; }
        reportChoices.Add(new ReportChoice {
            Kind = "PBIR",
            DefinitionFile = f,
            ReportFolder = Path.GetDirectoryName(f),
            ReportName = CleanReportName(Path.GetDirectoryName(f)),
            ConnectionSummary = ReadConnectionSummary(f),
            Selected = true
        });
    }
}

if (scanApps)
{
    foreach (var f in EnumerateSourceFiles(appsFolder).Where(p =>
        string.Equals(Path.GetFileName(p), "rayfin.yml", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(Path.GetFileName(Path.GetDirectoryName(p)), "rayfin", StringComparison.OrdinalIgnoreCase)))
    {
        var appRoot = Path.GetDirectoryName(Path.GetDirectoryName(f));
        // rayfin.yml connectors plus fabric.yaml semantic models (Data App template)
        var connectors = ReadAppConnectors(f).Concat(ReadFabricYaml(appRoot)).ToList();
        var matching = appModelFilter.Length == 0
            ? connectors
            : connectors.Where(c => c.Any(v => string.Equals(v, appModelFilter, StringComparison.OrdinalIgnoreCase))).ToList();
        // Without a filter every app is scanned, also one whose connection is not found.
        if (appModelFilter.Length > 0 && matching.Count == 0) { appsNotMatchingFilter++; continue; }
        reportChoices.Add(new ReportChoice {
            Kind = "App",
            DefinitionFile = f,
            ReportFolder = appRoot,
            ReportName = Path.GetFileName(appRoot),
            ConnectionSummary = matching.Count == 0
                ? "no semantic-model connection found"
                : string.Join(", ", matching.Select(c => c[0] + " (model " + (c[2] == "" ? "?" : c[2]) + ")")),
            Selected = true
        });
    }
}

reportChoices = reportChoices.OrderBy(r => r.Kind).ThenBy(r => r.ReportName).ThenBy(r => r.ReportFolder).ToList();

if (reportChoices.Count == 0)
{
    Info("Nothing to scan was found.\n\n" +
        (scanPbir ? "PBIR folder: " + pbirFolder + (reportsNotForThisModel > 0 ? "\n  " + reportsNotForThisModel + " report(s) skipped because they point to another model." : "\n  no definition.pbir found.") + "\n" : "") +
        (scanApps ? "Apps folder: " + appsFolder + (appsNotMatchingFilter > 0 ? "\n  " + appsNotMatchingFilter + " app(s) skipped: no connector matches the model filter \"" + appModelFilter + "\"." : "\n  no rayfin/rayfin.yml found.") + "\n" : "") +
        "\nHold Shift when starting the macro to change the folders.");
    return;
}

// ---------------------------------------------------------------------------
// Dialog: select reports and apps (skipped when starting straight away)
// ---------------------------------------------------------------------------
if (!autoRun)
using (var form = new Form())
{
    form.Text = "Select reports and apps to scan";
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
        Text = "[PBIR] rows are reports, [App] rows are Fabric Apps. Files must be available on disk; model dependencies are taken from the model currently open in Tabular Editor.",
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
            MessageBox.Show("Select at least one report or app.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

var selectedReports = reportChoices.Where(r => r.Selected && r.Kind == "PBIR").ToList();
var selectedApps = reportChoices.Where(r => r.Selected && r.Kind == "App").ToList();

// ---------------------------------------------------------------------------
// Dialog: choose object scope
// ---------------------------------------------------------------------------
bool includeMeasures = true;
bool includeColumns = true;
bool includeFunctions = true;
bool includeSupportTables = true;
bool selectedOnly = false;

var selectedMeasures = Selected.Measures.ToList();
var selectedColumns = Selected.Columns.ToList();
var selectedFunctions = Selected.Functions.ToList();
var selectedSupportTables = Selected.Tables.Where(t => IsExcludedSupportTable(t)).ToList();
bool hasSelectedObjects = selectedMeasures.Count > 0 || selectedColumns.Count > 0 || selectedFunctions.Count > 0 || selectedSupportTables.Count > 0;

using (var form = new Form())
{
    form.Text = "Choose model object scope";
    form.Width = 760;
    form.Height = 390;
    form.MinimumSize = new Size(620, 380);
    form.StartPosition = FormStartPosition.CenterScreen;
    form.FormBorderStyle = FormBorderStyle.Sizable;
    form.MaximizeBox = true;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 7, ColumnCount = 1, Padding = new Padding(12) };
    for (int i = 0; i < 6; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var info = new Label {
        Text = "Select what to classify. If you choose selected objects, only currently selected measures/columns in Tabular Editor are listed. PBIR visual references are still scanned from the selected reports.",
        Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(640, 0), Font = font
    };
    var chkMeasures = new CheckBox { Text = "Classify measures", Checked = true, AutoSize = true, Font = font };
    var chkColumns = new CheckBox { Text = "Classify columns", Checked = true, AutoSize = true, Font = font };
    var chkFunctions = new CheckBox { Text = "Classify DAX UDFs", Checked = true, AutoSize = true, Font = font };
    var chkSupportTables = new CheckBox { Text = "Classify field parameters and calculation groups (one row per table)", Checked = true, AutoSize = true, Font = font };
    var chkSelected = new CheckBox {
        Text = "Use selected objects only (" + selectedMeasures.Count + " measures, " + selectedColumns.Count + " columns, " + selectedFunctions.Count + " DAX UDFs, " +
            selectedSupportTables.Count + " field parameters / calculation groups selected)",
        Checked = hasSelectedObjects, Enabled = hasSelectedObjects, AutoSize = true, Font = font
    };
    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 12, 0, 0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.Add(ok);
    buttons.Controls.Add(cancel);
    ok.Click += (s, e) =>
    {
        if (!chkMeasures.Checked && !chkColumns.Checked && !chkFunctions.Checked && !chkSupportTables.Checked)
        {
            MessageBox.Show("Select measures, columns, DAX UDFs, field parameters / calculation groups, or a combination.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        includeMeasures = chkMeasures.Checked;
        includeColumns = chkColumns.Checked;
        includeFunctions = chkFunctions.Checked;
        includeSupportTables = chkSupportTables.Checked;
        selectedOnly = chkSelected.Checked;
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    layout.Controls.Add(info, 0, 0);
    layout.Controls.Add(chkMeasures, 0, 1);
    layout.Controls.Add(chkColumns, 0, 2);
    layout.Controls.Add(chkFunctions, 0, 3);
    layout.Controls.Add(chkSupportTables, 0, 4);
    layout.Controls.Add(chkSelected, 0, 5);
    layout.Controls.Add(buttons, 0, 6);
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

// Hierarchy levels used in visuals: "Table|Hierarchy|Level", mapped to the level's
// column after the scan
var visualHierarchyLevelKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

Func<JToken, Dictionary<string, string>, string> ResolveEntity = (sourceRef, aliases) =>
{
    var entity = (string)sourceRef?["Entity"];
    if (entity == null)
    {
        var source = (string)sourceRef?["Source"];
        if (source != null) aliases.TryGetValue(source, out entity);
    }
    return entity;
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

        // { "HierarchyLevel": { "Expression": { "Hierarchy": { "Expression": { "SourceRef": ... },
        //   "Hierarchy": "Calendar" } }, "Level": "Year" } }
        // Auto date/time hierarchies use PropertyVariationSource instead of SourceRef;
        // the Property is the date column itself.
        if (parentKey == "HierarchyLevel")
        {
            var hierarchy = obj["Expression"]?["Hierarchy"];
            var hierarchyName = (string)hierarchy?["Hierarchy"];
            var level = (string)obj["Level"];
            var table = ResolveEntity(hierarchy?["Expression"]?["SourceRef"], localAliases);
            if (!string.IsNullOrEmpty(table) && !string.IsNullOrEmpty(hierarchyName) && !string.IsNullOrEmpty(level))
                visualHierarchyLevelKeys.Add(table + "|" + hierarchyName + "|" + level);

            var variation = hierarchy?["Expression"]?["PropertyVariationSource"];
            var variationTable = ResolveEntity(variation?["Expression"]?["SourceRef"], localAliases);
            var variationColumn = (string)variation?["Property"];
            if (!string.IsNullOrEmpty(variationTable) && !string.IsNullOrEmpty(variationColumn))
            {
                visualColumnKeys.Add(variationTable + "|" + variationColumn);
                addCount("Column|" + variationTable + "|" + variationColumn);
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

// A hierarchy level used in a visual is visual use of the level's column
var hierarchyLevelUse = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
foreach (var t in Model.Tables)
    foreach (var h in t.Hierarchies)
        foreach (var l in h.Levels)
        {
            if (l.Column == null || !visualHierarchyLevelKeys.Contains(t.Name + "|" + h.Name + "|" + l.Name)) continue;
            var key = t.Name + "|" + l.Column.Name;
            visualColumnKeys.Add(key);
            addCount("Column|" + key);
            if (!hierarchyLevelUse.ContainsKey(key)) hierarchyLevelUse[key] = "'" + t.Name + "'[" + h.Name + "] level " + l.Name;
        }

// ---------------------------------------------------------------------------
// Scan Fabric App references (DAX in .dax files, model names in .ts/.tsx)
// ---------------------------------------------------------------------------
// App references count as direct use, exactly like PBIR visual references, so
// they go into the same visual*Keys sets. The PBIR-only sets are kept to word
// the Reason column.
var pbirMeasureKeys = new HashSet<string>(visualMeasureKeys, StringComparer.OrdinalIgnoreCase);
var pbirColumnKeys = new HashSet<string>(visualColumnKeys, StringComparer.OrdinalIgnoreCase);
var appMeasureKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var appColumnKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var appReferenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
var appNotInModel = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
int scannedAppFiles = 0;

var StringLiteral = new System.Text.RegularExpressions.Regex("\"((?:[^\"]|\"\")*)\"");
var QualifiedRef = new System.Text.RegularExpressions.Regex(@"(?<![\w\].])('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]");
var BareRef = new System.Text.RegularExpressions.Regex(@"\[((?:[^\]]|\]\])+)\]");
var DefineMeasure = new System.Text.RegularExpressions.Regex(@"\bMEASURE\s+('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]\s*=",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

Func<string, string> UnquoteTable = t => t.StartsWith("'") ? t.Substring(1, t.Length - 2).Replace("''", "'") : t;
Func<string, string> UnescapeName = n => n.Replace("]]", "]");

// Removes DAX comments and string literals; keeps 'quoted tables' and [names] intact.
Func<string, string> StripDaxCommentsAndStrings = dax =>
{
    var sb = new System.Text.StringBuilder(dax.Length);
    int i = 0;
    while (i < dax.Length)
    {
        char c = dax[i];
        char next = i + 1 < dax.Length ? dax[i + 1] : '\0';
        if ((c == '/' && next == '/') || (c == '-' && next == '-'))
        {
            while (i < dax.Length && dax[i] != '\n') i++;
            continue;
        }
        if (c == '/' && next == '*')
        {
            int end = dax.IndexOf("*/", i + 2);
            i = end < 0 ? dax.Length : end + 2;
            sb.Append(' ');
            continue;
        }
        if (c == '"')
        {
            i++;
            while (i < dax.Length)
            {
                if (dax[i] == '"' && i + 1 < dax.Length && dax[i + 1] == '"') { i += 2; continue; }
                if (dax[i] == '"') { i++; break; }
                i++;
            }
            sb.Append("\"\"");
            continue;
        }
        if (c == '\'' || c == '[')
        {
            char close = c == '\'' ? '\'' : ']';
            int start = i;
            i++;
            while (i < dax.Length)
            {
                if (dax[i] == close && i + 1 < dax.Length && dax[i + 1] == close) { i += 2; continue; }
                if (dax[i] == close) { i++; break; }
                i++;
            }
            sb.Append(dax, start, i - start);
            continue;
        }
        sb.Append(c);
        i++;
    }
    return sb.ToString();
};

var appTablesByName = Model.Tables.ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);
var appMeasuresByName = new Dictionary<string, Measure>(StringComparer.OrdinalIgnoreCase);
foreach (var am in Model.AllMeasures)
    if (!appMeasuresByName.ContainsKey(am.Name)) appMeasuresByName.Add(am.Name, am);

Func<string, string, Column> FindAppColumn = (table, name) =>
{
    Table t;
    if (!appTablesByName.TryGetValue(table, out t)) return null;
    return t.Columns.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
};
Action<Measure> addAppMeasure = am =>
{
    var key = am.Table.Name + "|" + am.Name;
    visualMeasureKeys.Add(key);
    appMeasureKeys.Add(key);
    int n; appReferenceCounts.TryGetValue("Measure|" + key, out n); appReferenceCounts["Measure|" + key] = n + 1;
};
Action<Column> addAppColumn = ac =>
{
    var key = ac.Table.Name + "|" + ac.Name;
    visualColumnKeys.Add(key);
    appColumnKeys.Add(key);
    int n; appReferenceCounts.TryGetValue("Column|" + key, out n); appReferenceCounts["Column|" + key] = n + 1;
};

foreach (var app in selectedApps)
{
    var appFiles = EnumerateSourceFiles(app.ReportFolder).ToList();

    foreach (var file in appFiles.Where(p => p.EndsWith(".dax", StringComparison.OrdinalIgnoreCase)))
    {
        scannedAppFiles++;
        var rel = app.ReportName + "/" + file.Substring(app.ReportFolder.Length).TrimStart('\\', '/').Replace('\\', '/');
        var rawDax = File.ReadAllText(file);
        var dax = StripDaxCommentsAndStrings(rawDax);

        // Result-column names created in the query ("Week" in SELECTCOLUMNS / ROW / ...)
        var queryAliases = new HashSet<string>(
            StringLiteral.Matches(rawDax).Cast<System.Text.RegularExpressions.Match>().Select(x => x.Groups[1].Value.Replace("\"\"", "\"")),
            StringComparer.OrdinalIgnoreCase);
        var queryMeasures = new HashSet<string>(
            DefineMeasure.Matches(dax).Cast<System.Text.RegularExpressions.Match>().Select(x => UnescapeName(x.Groups[2].Value)),
            StringComparer.OrdinalIgnoreCase);

        var remaining = new System.Text.StringBuilder(dax);
        foreach (System.Text.RegularExpressions.Match m in QualifiedRef.Matches(dax))
        {
            var table = UnquoteTable(m.Groups[1].Value);
            if (!appTablesByName.ContainsKey(table)) continue; // keyword (ORDER BY [x]) or VAR: bare pass
            var name = UnescapeName(m.Groups[2].Value);
            for (int k = m.Index; k < m.Index + m.Length; k++) remaining[k] = ' ';
            if (queryMeasures.Contains(name)) continue;

            var col = FindAppColumn(table, name);
            Measure meas;
            if (col != null) addAppColumn(col);
            else if (appMeasuresByName.TryGetValue(name, out meas)) addAppMeasure(meas);
            else appNotInModel.Add(table + "[" + name + "]  (" + rel + ")");
        }

        foreach (System.Text.RegularExpressions.Match m in BareRef.Matches(remaining.ToString()))
        {
            var name = UnescapeName(m.Groups[1].Value);
            Measure meas;
            if (appMeasuresByName.TryGetValue(name, out meas)) addAppMeasure(meas);
            else if (!queryMeasures.Contains(name) && !queryAliases.Contains(name))
                appNotInModel.Add("[" + name + "]  (" + rel + ")");
        }
    }

    // .ts/.tsx: only names that exist in the model (DAX fragments built in code,
    // measure names handed to DAX templates). Tests and type declarations skipped.
    foreach (var file in appFiles.Where(p =>
        (p.EndsWith(".ts", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".tsx", StringComparison.OrdinalIgnoreCase))
        && !System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(p), @"\.(spec|test)\.tsx?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        && !p.EndsWith(".d.ts", StringComparison.OrdinalIgnoreCase)))
    {
        scannedAppFiles++;
        var code = File.ReadAllText(file);
        foreach (System.Text.RegularExpressions.Match m in BareRef.Matches(code))
        {
            Measure meas;
            if (appMeasuresByName.TryGetValue(UnescapeName(m.Groups[1].Value), out meas)) addAppMeasure(meas);
        }
        foreach (System.Text.RegularExpressions.Match m in QualifiedRef.Matches(code))
        {
            var col = FindAppColumn(UnquoteTable(m.Groups[1].Value), UnescapeName(m.Groups[2].Value));
            if (col != null) addAppColumn(col);
        }
    }
}

// Reason text for direct use, naming where the use comes from.
Func<string, bool, string> DirectUseReason = (key, isMeasure) =>
{
    bool inPbir = (isMeasure ? pbirMeasureKeys : pbirColumnKeys).Contains(key);
    bool inApp = (isMeasure ? appMeasureKeys : appColumnKeys).Contains(key);
    if (inPbir && inApp) return "Referenced directly by selected PBIR report visuals/filters and queried by selected Fabric Apps.";
    if (inApp) return "Queried directly by selected Fabric Apps (DAX in the app).";
    return "Referenced directly by selected PBIR report visuals/filters.";
};

// ---------------------------------------------------------------------------
// Field parameters, calculation groups and perspectives
// ---------------------------------------------------------------------------
// DependsOn as a flat list of measures/columns/tables, robust to the collection shape
Func<object, List<object>> DependencyObjects = o =>
{
    var list = new List<object>();
    var prop = o == null ? null : o.GetType().GetProperty("DependsOn");
    var deps = prop == null ? null : prop.GetValue(o, null);
    if (deps == null) return list;
    foreach (var name in new[] { "Measures", "Columns", "Tables" })
    {
        var sub = deps.GetType().GetProperty(name);
        var items = sub == null ? null : sub.GetValue(deps, null) as System.Collections.IEnumerable;
        if (items != null) list.AddRange(items.Cast<object>());
    }
    var all = deps as System.Collections.IEnumerable;
    if (all != null)
        foreach (var item in all)
        {
            var keyProp = item.GetType().GetProperty("Key");
            list.Add(keyProp != null ? keyProp.GetValue(item, null) : item);
        }
    return list.Where(x => x != null).Distinct().ToList();
};

Func<Table, string> SupportTableKind = t => t is CalculationGroupTable ? "Calculation group" : "Field parameter";
Func<Table, bool> IsUsedInReports = t => t.Columns.Any(c => visualColumnKeys.Contains(t.Name + "|" + c.Name));
var supportTables = Model.Tables.Where(t => IsExcludedSupportTable(t)).ToList();

// Measures/columns listed in a field parameter used in the reports -> shown in visuals
// Measures/columns referenced by a calculation item of a calculation group used in the reports
var shownByFieldParameter = new Dictionary<object, string>();
var usedByCalculationItem = new Dictionary<object, string>();
foreach (var t in supportTables.Where(IsUsedInReports))
{
    var cg = t as CalculationGroupTable;
    if (cg == null)
    {
        foreach (var dep in DependencyObjects(t).Where(d => d is Measure || d is Column))
            if (!shownByFieldParameter.ContainsKey(dep)) shownByFieldParameter[dep] = t.Name;
    }
    else
    {
        foreach (var item in cg.CalculationItems)
            foreach (var dep in DependencyObjects(item).Where(d => d is Measure || d is Column || d is Function))
                if (!usedByCalculationItem.ContainsKey(dep)) usedByCalculationItem[dep] = "'" + cg.Name + "'[" + item.Name + "]";
    }
}

// Perspective membership (measures, columns, tables)
var perspectiveNames = Model.Perspectives.Select(p => p.Name).ToList();
Func<object, string> PerspectivesOf = o =>
{
    var names = new List<string>();
    try
    {
        foreach (var p in perspectiveNames)
        {
            var m = o as Measure; var c = o as Column; var t = o as Table;
            if ((m != null && m.InPerspective[p]) || (c != null && c.InPerspective[p]) || (t != null && t.InPerspective[p])) names.Add(p);
        }
    }
    catch { }
    return string.Join(", ", names);
};

Func<object, string> DescribeDependent = d =>
{
    var cc = d as CalculatedColumn;
    if (cc != null) return "calculated column '" + cc.Table.Name + "'[" + cc.Name + "]";
    var tbl = d as Table;
    if (tbl != null) return "calculated table '" + tbl.Name + "'";
    var tp = d as TablePermission;
    if (tp != null) return "RLS filter of role '" + tp.Role.Name + "' on '" + tp.Table.Name + "'";
    var ci = d as CalculationItem;
    if (ci != null) return "calculation item [" + ci.Name + "]";
    var named = d as ITabularNamedObject;
    return d.GetType().Name + (named != null ? " " + named.Name : "");
};

Func<IEnumerable<string>, string> NameList = items =>
{
    var list = items.Distinct().ToList();
    return string.Join(", ", list.Take(3)) + (list.Count > 3 ? " (+" + (list.Count - 3) + " more)" : "");
};

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
    string via;
    if (visualMeasureKeys.Contains(key))
    {
        measureCategory[m] = CatVisual;
        measureReasons[m] = DirectUseReason(key, true);
    }
    else if (shownByFieldParameter.TryGetValue(m, out via))
    {
        measureCategory[m] = CatVisual;
        measureReasons[m] = "Shown in visuals through field parameter '" + via + "', which is used in the selected PBIR reports / Fabric Apps.";
    }
    else if (usedByCalculationItem.TryGetValue(m, out via))
    {
        measureCategory[m] = CatMeasureWithVisual;
        measureReasons[m] = "Used by calculation item " + via + " of a calculation group used in the selected PBIR reports / Fabric Apps.";
    }
}

foreach (var f in functionClassificationCandidates)
{
    string via;
    if (usedByCalculationItem.TryGetValue(f, out via))
    {
        functionCategory[f] = CatDaxUdfWithVisual;
        functionReasons[f] = "Used by calculation item " + via + " of a calculation group used in the selected PBIR reports / Fabric Apps.";
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

// Perspectives: a measure in a perspective - or used (directly, or through other
// measures / DAX UDFs) by one - is a model dependency for downstream use
{
    var queue = new Queue<object>();
    foreach (var m in measureClassificationCandidates)
    {
        if (measureCategory.ContainsKey(m)) continue;
        var names = PerspectivesOf(m);
        if (names.Length == 0) continue;
        measureCategory[m] = CatPerspective;
        measureReasons[m] = "In perspective(s): " + names + " - model dependency, no selected PBIR visual dependency.";
        queue.Enqueue(m);
    }
    var visited = new HashSet<object>(queue);
    while (queue.Count > 0)
    {
        var current = queue.Dequeue();
        foreach (var dep in DependencyObjects(current))
        {
            if (!(dep is Measure) && !(dep is Function)) continue;
            if (!visited.Add(dep)) continue;
            var depMeasure = dep as Measure;
            if (depMeasure != null && !measureCategory.ContainsKey(depMeasure) && !IsExcludedSupportTable(depMeasure.Table))
            {
                var currentMeasure = current as Measure;
                measureCategory[depMeasure] = CatPerspective;
                measureReasons[depMeasure] = (currentMeasure != null
                    ? "Used by [" + currentMeasure.Name + "], which is in (or used by) a perspective"
                    : "Used by a DAX UDF that a measure in a perspective uses") +
                    " - model dependency, no selected PBIR visual dependency.";
            }
            var depFunction = dep as Function;
            if (depFunction != null && !functionCategory.ContainsKey(depFunction))
            {
                functionCategory[depFunction] = CatPerspective;
                functionReasons[depFunction] = "Used by a measure in (or used by) a perspective - model dependency, no selected PBIR visual dependency.";
            }
            queue.Enqueue(dep);
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
    // Calculation items (of calculation groups not used in the reports), calculated
    // columns / tables and RLS filters calling the function
    var modelDependents = GetReferencedBy(f).Where(d => !(d is Measure) && !(d is Function)).ToList();
    if (modelDependents.Any())
    {
        functionCategory[f] = CatModelNoVisual;
        functionReasons[f] = "Model dependency without visual dependency: " + NameList(modelDependents.Select(DescribeDependent)) +
            (hasDaxUdfDependents || hasMeasureDependents ? ". Also used by DAX UDFs/measures without visual dependency." : ".");
    }
    else if (hasDaxUdfDependents || hasMeasureDependents)
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
//   Ranking, highest first:
//     VisualDependency                                     used directly in visuals / field parameter
//     ColumnUsedByVisualDependentMeasure                   needed by a measure, DAX UDF, calculation item,
//                                                          calculated column/table or sort-by with visual use
//     PerspectiveDependencyWithoutVisualDependency         in a perspective, or used by a perspective measure
//     ModelDependencyWithoutVisualDependency               relationship, hierarchy, sort-by, Calendar,
//                                                          calculated column/table, RLS, calculation item
//     ColumnUsedByDaxUdfDependency                         only DAX UDFs (and measures) without visual use
//     ColumnUsedByMeasureDependencyWithoutVisualDependency only measures without visual use
//     NoDependencies
//   Calculated columns and calculated table columns are classified like data
//   columns (Model.AllColumns); only field parameter / calculation group
//   tables are handled separately.
// ---------------------------------------------------------------------------
var columnCategory = new Dictionary<Column, string>();
var columnReasons = new Dictionary<Column, string>();
// All columns are classified, so dependency chains also work with "selected objects only"
var allModelColumns = Model.AllColumns.Where(c => !IsExcludedSupportTable(c.Table)).ToList();

Func<Column, bool> HasRelationshipDependency = c =>
    Model.Relationships.Any(r => r.FromColumn == c || r.ToColumn == c);

Func<Column, bool> HasHierarchyDependency = c =>
    c.Table.Hierarchies.Any(h => h.Levels.Any(l => l.Column == c));

// Only columns sorted BY this column depend on it. The column this one is
// sorted by is what it needs, not what makes it used (same rule as DependsOn).
Func<Column, bool> HasSortByDependency = c =>
    c.Table.Columns.Any(other => other.SortByColumn == c);

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

Func<Measure, bool> MeasureHasVisual = m =>
    measureCategory.ContainsKey(m) && (measureCategory[m] == CatVisual || measureCategory[m] == CatMeasureWithVisual);
Func<Function, bool> FunctionHasVisual = f =>
    functionCategory.ContainsKey(f) && functionCategory[f] == CatDaxUdfWithVisual;

// DAX objects other than measures and DAX UDFs that reference the column:
// calculated columns, calculated tables, RLS filters, calculation items ...
Func<Column, List<object>> ModelDaxDependents = c =>
    GetReferencedBy(c).Where(d => !(d is Measure) && !(d is Function)).ToList();

// 1. Visual use: direct, through a field parameter, a used calculation group,
//    or a measure / DAX UDF with visual dependency
foreach (var c in allModelColumns)
{
    var key = c.Table.Name + "|" + c.Name;
    string via;
    if (visualColumnKeys.Contains(key))
    {
        columnCategory[c] = CatVisual;
        columnReasons[c] = DirectUseReason(key, false) +
            (hierarchyLevelUse.TryGetValue(key, out via) ? " Used through hierarchy " + via + "." : "");
        continue;
    }
    if (shownByFieldParameter.TryGetValue(c, out via))
    {
        columnCategory[c] = CatVisual;
        columnReasons[c] = "Shown in visuals through field parameter '" + via + "', which is used in the selected PBIR reports / Fabric Apps.";
        continue;
    }
    if (usedByCalculationItem.TryGetValue(c, out via))
    {
        columnCategory[c] = CatColumnUsedByVisualMeasure;
        columnReasons[c] = "Used by calculation item " + via + " of a calculation group used in the selected PBIR reports / Fabric Apps.";
        continue;
    }
    var visualMeasures = c.ReferencedBy.OfType<Measure>().Where(MeasureHasVisual).ToList();
    var visualFunctions = GetReferencedBy(c).OfType<Function>().Where(FunctionHasVisual).ToList();
    if (visualMeasures.Any())
    {
        columnCategory[c] = CatColumnUsedByVisualMeasure;
        columnReasons[c] = "Used by measure(s) with visual dependency: " + NameList(visualMeasures.Select(m => "[" + m.Name + "]")) + ".";
    }
    else if (visualFunctions.Any())
    {
        columnCategory[c] = CatColumnUsedByVisualMeasure;
        columnReasons[c] = "Used by DAX UDF(s) with visual dependency: " + NameList(visualFunctions.Select(f => f.Name)) + ".";
    }
}

// 2. Columns needed by a visually-used column: referenced by a calculated
//    column / calculated table with visual use, or the sort-by column of one
bool columnChanged = true;
while (columnChanged)
{
    columnChanged = false;
    foreach (var c in allModelColumns)
    {
        if (columnCategory.ContainsKey(c)) continue;
        string via = null;
        foreach (var d in ModelDaxDependents(c))
        {
            var depColumn = d as Column;
            var depTable = d as Table;
            if ((depColumn != null && columnCategory.ContainsKey(depColumn)) ||
                (depTable != null && depTable.Columns.Any(tc => columnCategory.ContainsKey(tc))))
            {
                via = DescribeDependent(d);
                break;
            }
        }
        if (via == null)
        {
            var sorted = c.Table.Columns.FirstOrDefault(o => o.SortByColumn == c && columnCategory.ContainsKey(o));
            if (sorted != null) via = "sort order of '" + c.Table.Name + "'[" + sorted.Name + "]";
        }
        if (via == null) continue;
        columnCategory[c] = CatColumnUsedByVisualMeasure;
        columnReasons[c] = "Needed by " + via + ", which has visual dependency.";
        columnChanged = true;
    }
}

// 3. No visual use: perspective > model dependency > DAX UDFs only > measures only > none
foreach (var c in allModelColumns)
{
    if (columnCategory.ContainsKey(c)) continue;

    var referencingMeasures = c.ReferencedBy.OfType<Measure>().ToList();
    var referencingFunctions = GetReferencedBy(c).OfType<Function>().ToList();
    var columnPerspectives = PerspectivesOf(c);
    bool usedByPerspectiveMeasure = referencingMeasures.Any(m =>
        measureCategory.ContainsKey(m) && measureCategory[m] == CatPerspective);

    var modelDependencies = new List<string>();
    if (HasRelationshipDependency(c)) modelDependencies.Add("relationship");
    modelDependencies.AddRange(c.Table.Hierarchies.Where(h => h.Levels.Any(l => l.Column == c)).Select(h => "hierarchy level in [" + h.Name + "]"));
    if (HasSortByDependency(c)) modelDependencies.Add("sort-by column of another column");
    if (HasCalendarOptionDependency(c)) modelDependencies.Add("Calendar");
    modelDependencies.AddRange(ModelDaxDependents(c).Select(DescribeDependent));

    bool usedByDaxUdf = referencingFunctions.Any() || referencingMeasures.Any(m =>
        measureCategory.ContainsKey(m) && measureCategory[m] == CatDaxUdfNoVisual);

    if (columnPerspectives.Length > 0 || usedByPerspectiveMeasure)
    {
        columnCategory[c] = CatPerspective;
        columnReasons[c] = (columnPerspectives.Length > 0
            ? "In perspective(s): " + columnPerspectives
            : "Used by measures in (or used by) a perspective") + " - model dependency, no selected PBIR visual dependency.";
    }
    else if (modelDependencies.Count > 0)
    {
        columnCategory[c] = CatModelNoVisual;
        columnReasons[c] = "Model dependency without visual dependency: " + NameList(modelDependencies) +
            (referencingMeasures.Any() ? ". Also used by measure(s) without visual dependency." : ".");
    }
    else if (usedByDaxUdf)
    {
        columnCategory[c] = CatColumnUsedByDaxUdfMeasure;
        columnReasons[c] = "Only used by DAX UDFs / measures without visual dependency (no relationship, hierarchy or other model dependency).";
    }
    else if (referencingMeasures.Any())
    {
        columnCategory[c] = CatColumnUsedByNoVisualMeasure;
        columnReasons[c] = "Only used by measure(s) without visual dependency: " + NameList(referencingMeasures.Select(m => "[" + m.Name + "]")) +
            " (no relationship, hierarchy or other model dependency).";
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
if (includeSupportTables)
{
    // Field parameters and calculation groups: used in the reports, or only a model dependency
    foreach (var t in selectedOnly ? selectedSupportTables : supportTables)
    {
        var tableColumns = t.Columns.ToList();
        int refs = tableColumns.Sum(c => { int n; visualReferenceCounts.TryGetValue("Column|" + t.Name + "|" + c.Name, out n); return n; });
        var referencingMeasures = tableColumns.SelectMany(c => c.ReferencedBy.OfType<Measure>()).Distinct().ToList();
        var perspectives = PerspectivesOf(t);
        string category, reason;
        if (IsUsedInReports(t))
        {
            category = CatVisual;
            reason = SupportTableKind(t) + " used directly by selected PBIR report visuals/filters.";
        }
        else if (referencingMeasures.Any(m => measureCategory.ContainsKey(m) && (measureCategory[m] == CatVisual || measureCategory[m] == CatMeasureWithVisual)))
        {
            category = CatColumnUsedByVisualMeasure;
            reason = "Not in the selected PBIR reports / Fabric Apps, but used by measures with visual dependency (e.g. " +
                string.Join(", ", referencingMeasures.Select(m => "[" + m.Name + "]").Take(3)) + ").";
        }
        else if (perspectives.Length > 0)
        {
            category = CatPerspective;
            reason = "In perspective(s): " + perspectives + " - model dependency, no selected PBIR visual dependency.";
        }
        else if (referencingMeasures.Any() || tableColumns.Any(c => c.ReferencedBy.Any()))
        {
            category = CatModelNoVisual;
            reason = "Referenced by model objects, but not used in the selected PBIR reports / Fabric Apps.";
        }
        else
        {
            category = CatNone;
            reason = "Not used in the selected PBIR reports and not referenced in the model.";
        }
        results.Add(new ObjectResult {
            Kind = SupportTableKind(t),
            Table = t.Name,
            Name = t.Name,
            Category = category,
            Reason = reason,
            Hidden = t.IsHidden,
            VisualReferences = refs,
            DependsOnCount = DependencyObjects(t).Count + (t is CalculationGroupTable ? ((CalculationGroupTable)t).CalculationItems.Sum(i => DependencyObjects(i).Count) : 0),
            ReferencedByCount = referencingMeasures.Count,
            TableObject = t
        });
    }
}
foreach (var r in results)
    r.Perspectives = PerspectivesOf((object)r.MeasureObject ?? (object)r.ColumnObject ?? r.TableObject);

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
Func<ObjectResult, int> AppRefs = r =>
{
    int n;
    return appReferenceCounts.TryGetValue(r.Kind + "|" + r.Table + "|" + r.Name, out n) ? n : 0;
};
var dt = new System.Data.DataTable();
dt.Columns.Add("Kind");
dt.Columns.Add("Table");
dt.Columns.Add("Name");
dt.Columns.Add("DependencyCategory");
dt.Columns.Add("Duplicate / redundant");
dt.Columns.Add("Display folder");
dt.Columns.Add("Hidden");
dt.Columns.Add("Perspectives");
dt.Columns.Add("PBIR refs");
dt.Columns.Add("App refs");
dt.Columns.Add("DependsOn (uses)");
dt.Columns.Add("ReferencedBy (used by)");
dt.Columns.Add("Reason");
foreach (var r in results.OrderBy(r => r.Kind).ThenBy(r => r.Category).ThenBy(r => r.Table).ThenBy(r => r.Name))
    dt.Rows.Add(
        r.Kind, r.Table, r.Name, r.Category, DuplicateOf(r),
        r.MeasureObject != null ? (r.MeasureObject.DisplayFolder ?? "") : r.ColumnObject != null ? (r.ColumnObject.DisplayFolder ?? "") : "",
        r.Hidden ? "Yes" : "No", r.Perspectives ?? "", r.VisualReferences, AppRefs(r), r.DependsOnCount, r.ReferencedByCount, r.Reason);
Output(dt);

// The output grid is the last step: this macro is read-only. Organizing
// (display folders, hiding, deleting) is done by the dedicated macros, e.g. 19.
Info(
    "Done - read-only, nothing in the model was changed.\n\n" +
    "Reports included: " + selectedReports.Count + "  |  PBIR JSON files scanned: " + scannedJsonFiles +
    (skippedJsonFiles > 0 ? " (" + skippedJsonFiles + " skipped)" : "") + "\n" +
    "Fabric Apps included: " + selectedApps.Count + "  |  App files scanned: " + scannedAppFiles + "\n" +
    (appNotInModel.Count > 0 ? "App references not found in the model (" + appNotInModel.Count + "):\n  " + string.Join("\n  ", appNotInModel.Take(10)) + (appNotInModel.Count > 10 ? "\n  ..." : "") + "\n" : "") + "\n" +
    string.Join("\n", results.GroupBy(r => r.Kind + " - " + r.Category).OrderBy(grp => grp.Key).Select(grp => grp.Key + ": " + grp.Count())) + "\n\n" +
    "Duplicate / redundant objects flagged: " + results.Count(r => DuplicateOf(r) != "") + " (see the output grid)");

