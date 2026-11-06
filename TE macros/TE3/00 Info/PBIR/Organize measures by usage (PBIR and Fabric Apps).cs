// ============================================================================
// 19 - Organize measures by usage (PBIR and Fabric Apps).cs
//
// THE macro for organizing measures. Organizes measures into display folders
// by how they are used in PBIR reports and Fabric Apps, so it is easy to see
// what is in use (and to document only the measures that are in use). Changes
// are made in the model and can be undone with Ctrl+Z in Tabular Editor. PBIR
// and app files are only read. (Macro 22 is the read-only overview of the same
// usage, also for columns, DAX UDFs and support tables.)
//
// Workflow:
//   1. Choose the sources: PBIR reports and/or Fabric Apps (rayfin/rayfin.yml),
//      each with a folder - same dialog and remembered settings as macro 22
//      (defaults to the repository root above the open model; reports can be
//      limited to the ones pointing to this model, apps by a model filter).
//      Then select which reports and apps to include. With "start straight
//      away", both dialogs are skipped next time (hold Shift to show them).
//   2. The reports are scanned for measure usage: visual wells (incl.
//      tooltips), sort, visual filters, formatting (conditional formatting vs
//      dynamic text), field parameters and calculation groups used in
//      visuals. Fabric Apps: a measure queried in a .dax file (or named in
//      .ts/.tsx code) counts as shown as a value, i.e. KPI. Perspective
//      membership is read from the model.
//   3. One dialog to set it all up:
//        - Scope: all measures, or only the measures selected in TE
//        - Categories, checked from top to bottom (first match wins):
//            Conditional formatting  used in visuals only for formatting
//                                    (colors, fills, icons, data bars, ...)
//            Dynamic text            used in visuals only as text (titles,
//                                    subtitles, labels, button/text values)
//            SVG / images            DataCategory ImageUrl and actually used in
//                                    a visual (otherwise it falls through to
//                                    Helper measures / No visual dependency)
//          Detection is automatic from the reports. A measure that is also
//          shown as a value (well, tooltip, sort, card reference label, ...)
//          is a KPI. "Name contains" is optional for a manual override.
//            Visual filter           used only as a visual-level filter
//            KPI (used in visuals)   in a visual well, sort, tooltip, via a
//                                    field parameter, or shown as a value in
//                                    formatting (card reference labels, error
//                                    bars, reference lines), or queried by a
//                                    Fabric App
//            Helper measures         referenced (directly or indirectly) by a
//                                    used measure, a DAX UDF, or a calculation
//                                    item of a calculation group used in visuals
//                                    (optionally hidden)
//            In perspective          in a perspective, or used by a measure in
//                                    a perspective - a model dependency for
//                                    downstream use, even without visual usage
//            No visual dependency    everything else (always last); duplicate
//                                    measures (same DAX as another measure)
//                                    optionally in a "Duplicates" subfolder
//          Per category: apply on/off, folder name, "name contains" text
//          (separate several with ;), order (Move up/down), and your own
//          categories (name-based).
//        - Folder placement:
//            Root + keep existing folders   KPI\Sales\Margin   (default)
//            Root only                      KPI  (original folder is kept
//                                           in an annotation for later)
//            Subfolder of existing folder   Sales\Margin\KPI
//        - Annotation MeasureUsage = <category> for model documentation.
//        - Optional: hide the helper measures (off by default).
//   4. Preview with a filter box, then apply.
//
// Re-running is safe: each measure remembers its own folder (annotation
// MeasureUsageOriginalFolder) and the folder this macro gave it
// (MeasureUsageAppliedFolder). If the measure is still in that folder, the
// next run starts from the remembered folder - so renaming KPI to
// VisualMeasures\KPI gives VisualMeasures\KPI\Sales, and a folder of your own
// with the same name as a category is kept. Measures moved by hand since are
// cleaned up by removing category folder names (also earlier names) and
// the "XX - Not to be used" folder of older versions of macro 16 / 22 from the
// start and end of the path.
//
// Scope: only the model open in Tabular Editor and the PBIR reports / Fabric Apps
// in the selected folders are analysed - see macro 22 for what is not detected
// (reports elsewhere in the Service, Copilot / AI, Excel, notebooks, ...).
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

const string AnnUsage = "MeasureUsage";
const string AnnOriginalFolder = "MeasureUsageOriginalFolder";   // the measure's own folder, without category folders
const string AnnAppliedFolder = "MeasureUsageAppliedFolder";     // the folder this macro set last time
const string AnnMacro16OriginalFolder = "DependencyOriginalDisplayFolder";
const string AnnKnownFolders = "MeasureUsageFolders";     // model-level: folder names used before

class ReportChoice
{
    public string Kind = "PBIR";   // PBIR or App (Fabric App)
    public string ReportFolder;    // report folder, or the app root
    public string ReportName;
    public string ConnectionSummary;
    public bool Selected;
    public string Label { get { return "[" + Kind + "] " + ReportName + "  |  " + ConnectionSummary + "  |  " + ReportFolder; } }
}

class UsageSignals
{
    public bool Projection, Tooltip, Sort, Filter, ConditionalFormatting, Text, FieldParameter, App, Helper, Perspective;
    public int References;
    public HashSet<string> Details = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
}

// True when a measure was actually found in a visual (well, tooltip, sort, filter,
// conditional formatting, dynamic text or field parameter) or a Fabric App query -
// i.e. it has a real visual dependency, as opposed to only being inferred from a
// property like DataCategory
bool HasVisualUsage(UsageSignals sig)
{
    return sig.Projection || sig.Tooltip || sig.Sort || sig.Filter || sig.ConditionalFormatting || sig.Text || sig.FieldParameter || sig.App;
}

class Category
{
    public string Key;          // CF, Text, SVG, Filter, KPI, Helper, None, Custom
    public string Name;
    public string Folder;
    public string Contains;
    public bool Apply = true;
    public string DetectedBy;
}

class MeasurePlan
{
    public Measure Measure;
    public Category Category;
    public string Reason;
    public bool Duplicate;
    public string DuplicateOf;
    public string NewFolder;
    public string BaseFolder;
    public bool Hide;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
bool SameName(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

Size DialogSize(double w, double h)
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * w), (int)(wa.Height * h));
}

string CleanReportName(string folder)
{
    var name = Path.GetFileName(folder);
    return name != null && name.EndsWith(".Report", StringComparison.OrdinalIgnoreCase) ? name.Substring(0, name.Length - ".Report".Length) : name;
}

string ReadConnectionSummary(string definitionFile)
{
    try
    {
        var json = JObject.Parse(File.ReadAllText(definitionFile));
        var dataset = json["datasetReference"] ?? json["semanticModelReference"];
        if (!string.IsNullOrWhiteSpace((string)dataset?["byPath"]?["path"])) return "byPath: " + (string)dataset["byPath"]["path"];
        if (!string.IsNullOrWhiteSpace((string)dataset?["byConnection"]?["connectionString"])) return "byConnection/live";
    }
    catch { }
    return "connection unknown";
}

// DependsOn in the TE object model, robust to the collection shape
List<object> DependenciesOf(object o)
{
    var result = new List<object>();
    var prop = o.GetType().GetProperty("DependsOn");
    var deps = prop == null ? null : prop.GetValue(o, null);
    if (deps == null) return result;
    foreach (var name in new[] { "Measures", "Columns", "Tables" })
    {
        var sub = deps.GetType().GetProperty(name);
        var items = sub == null ? null : sub.GetValue(deps, null) as System.Collections.IEnumerable;
        if (items != null) result.AddRange(items.Cast<object>());
    }
    var all = deps as System.Collections.IEnumerable;
    if (all != null)
        foreach (var item in all)
        {
            var keyProp = item.GetType().GetProperty("Key");
            result.Add(keyProp != null ? keyProp.GetValue(item, null) : item);
        }
    return result.Where(x => x != null).Distinct().ToList();
}

string NormalizeDax(string expression)
{
    var noComments = Regex.Replace(expression ?? "", @"//[^\n]*|--[^\n]*|/\*.*?\*/", "", RegexOptions.Singleline);
    return Regex.Replace(noComments, @"\s+", "").ToUpperInvariant();
}

TableLayoutPanel CreateDialogLayout(Form dlg, string title, int rows, Size size)
{
    dlg.Text = title;
    dlg.Size = size;
    dlg.MinimumSize = new Size(760, 480);
    dlg.StartPosition = FormStartPosition.CenterScreen;
    dlg.FormBorderStyle = FormBorderStyle.Sizable;
    dlg.MaximizeBox = true;
    dlg.MinimizeBox = false;
    dlg.Font = new Font("Segoe UI", 9);
    var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = rows, ColumnCount = 1, Padding = new Padding(12) };
    dlg.Controls.Add(panel);
    return panel;
}
Label Heading(string text)
{
    return new Label { Text = text, AutoSize = true, MaximumSize = new Size(1200, 0), Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 6) };
}

var editBack = Color.FromArgb(255, 244, 179);
var editHeaderBack = Color.FromArgb(255, 214, 102);
var naBack = Color.FromArgb(235, 235, 235);
void Highlight(DataGridViewColumn column)
{
    column.DefaultCellStyle = new DataGridViewCellStyle {
        BackColor = editBack, ForeColor = Color.Black, Font = new Font("Segoe UI", 9, FontStyle.Bold),
        SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black
    };
    column.HeaderCell.Style.BackColor = editHeaderBack;
    column.HeaderCell.Style.ForeColor = Color.Black;
}

// ---------------------------------------------------------------------------
// Step 1: sources - PBIR reports and/or Fabric Apps (same dialog and remembered
// settings as macro 22, so both macros scan the same sources)
// ---------------------------------------------------------------------------
string[] SkippedFolders = { "node_modules", "dist", ".agents", ".git", ".temp", ".vite", "coverage", "bin", "obj" };

// Walks a folder tree without descending into skipped folders (node_modules alone
// can hold tens of thousands of files) or folders that cannot be read.
List<string> EnumerateSourceFiles(string root)
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
}

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
        var boldFont = new Font("Segoe UI", 9, FontStyle.Bold);
        form.Font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3, Padding = new Padding(14) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 480));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        Action<Control, int> addSpan = (c, span) => { layout.Controls.Add(c); layout.SetColumnSpan(c, span); };

        addSpan(new Label {
            Text = "Model: " + (modelName == "" ? "(unnamed)" : modelName) + "\n" +
                   (modelSourcePath != null ? "Loaded from: " + modelSourcePath : "Not loaded from a file - pick folders below; they are remembered for this model.") + "\n" +
                   "Include every report and app that uses the model - a measure used only in a source you leave out ends up as 'No visual dependency'.",
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

        var chkRemember = new CheckBox { Text = "Remember these settings for this model (shared with the dependency classification)", Checked = true, AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
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
bool PointsToOpenModel(string definitionFile)
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
}

// Semantic-model connectors from the top-level "connectors:" block of rayfin.yml
// (services.connectors is only the on/off switch). Each entry: alias, workspaceId, itemId.
List<string[]> ReadAppConnectors(string rayfinYml)
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
}

// Semantic models declared in the app's fabric.yaml (fabric-app-data, used by the
// Data App template) for the active profile: profiles.<p>.semanticModels.<alias>.
List<string[]> ReadFabricYaml(string appRoot)
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
}

var reportChoices = new List<ReportChoice>();
int reportsNotForThisModel = 0, appsNotMatchingFilter = 0;

if (scanPbir)
{
    foreach (var f in EnumerateSourceFiles(pbirFolder).Where(p => string.Equals(Path.GetFileName(p), "definition.pbir", StringComparison.OrdinalIgnoreCase)))
    {
        if (onlyThisModelReports && !PointsToOpenModel(f)) { reportsNotForThisModel++; continue; }
        reportChoices.Add(new ReportChoice {
            Kind = "PBIR",
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
    Info("Nothing to scan was found - no changes made.\n\n" +
        (scanPbir ? "PBIR folder: " + pbirFolder + (reportsNotForThisModel > 0 ? "\n  " + reportsNotForThisModel + " report(s) skipped because they point to another model." : "\n  no definition.pbir found.") + "\n" : "") +
        (scanApps ? "Apps folder: " + appsFolder + (appsNotMatchingFilter > 0 ? "\n  " + appsNotMatchingFilter + " app(s) skipped: no connector matches the model filter \"" + appModelFilter + "\"." : "\n  no rayfin/rayfin.yml found.") + "\n" : "") +
        "\nHold Shift when starting the macro to change the folders.");
    return;
}

// Select reports and apps (skipped when starting straight away)
if (!autoRun)
using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Select reports and apps to scan", 4, DialogSize(0.85, 0.85));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var searchBox = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 8) };
    var reportList = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, HorizontalScrollbar = true };
    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var checkAll = new Button { Text = "Check all visible", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    var clearAll = new Button { Text = "Clear all", AutoSize = true, Margin = new Padding(0, 0, 24, 0) };
    var okButton = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { checkAll, clearAll, okButton, cancelButton });

    bool suppress = false;
    Action rebuild = () =>
    {
        suppress = true;
        reportList.Items.Clear();
        var filterText = (searchBox.Text ?? "").Trim();
        foreach (var rc in reportChoices.Where(rc => filterText.Length == 0 || rc.Label.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0))
            reportList.Items.Add(rc.Label, rc.Selected);
        suppress = false;
    };
    reportList.ItemCheck += (s, e) =>
    {
        if (suppress) return;
        var itemLabel = (string)reportList.Items[e.Index];
        reportChoices.First(rc => rc.Label == itemLabel).Selected = e.NewValue == CheckState.Checked;
    };
    searchBox.TextChanged += (s, e) => rebuild();
    checkAll.Click += (s, e) =>
    {
        foreach (var rc in reportChoices.Where(rc => string.IsNullOrEmpty(searchBox.Text) || rc.Label.IndexOf(searchBox.Text, StringComparison.OrdinalIgnoreCase) >= 0)) rc.Selected = true;
        rebuild();
    };
    clearAll.Click += (s, e) => { foreach (var rc in reportChoices) rc.Selected = false; rebuild(); };
    okButton.Click += (s, e) =>
    {
        if (!reportChoices.Any(rc => rc.Selected)) { MessageBox.Show("Select at least one report or app.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(Heading("Select the reports ([PBIR]) and Fabric Apps ([App]) to scan for measure usage. Include every one that uses the model - a measure used only in a source you leave out ends up as 'No visual dependency'."), 0, 0);
    panel.Controls.Add(searchBox, 0, 1);
    panel.Controls.Add(reportList, 0, 2);
    panel.Controls.Add(buttonRow, 0, 3);
    dlg.AcceptButton = okButton;
    dlg.CancelButton = cancelButton;
    rebuild();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}
var selectedReports = reportChoices.Where(rc => rc.Selected && rc.Kind == "PBIR").ToList();
var selectedApps = reportChoices.Where(rc => rc.Selected && rc.Kind == "App").ToList();

// ---------------------------------------------------------------------------
// Step 2: scan PBIR for measure usage
// ---------------------------------------------------------------------------
var measuresByName = new Dictionary<string, Measure>(StringComparer.OrdinalIgnoreCase);
foreach (var ms in Model.AllMeasures) measuresByName[ms.Name] = ms;     // measure names are unique in a model
var signals = Model.AllMeasures.ToDictionary(ms => ms, ms => new UsageSignals());
var usedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
int scannedFiles = 0, skippedFiles = 0;

// A measure used in visual formatting is one of:
//   "CF"    styling - colors, fills, icons, data bars, sizes (conditional formatting), also on titles
//   "Text"  text shown to the user - title/subtitle text, button/text values, dynamic labels
//   "Value" a value shown in the visual - card reference labels, error bars, reference lines
string FormattingKind(string propertyName)
{
    var p = (propertyName ?? "").ToLowerInvariant();
    var styling = new[] { "color", "fill", "background", "border", "icon", "bar", "shadow", "transparency", "size", "weight", "font", "visible", "show" };
    if (styling.Any(x => p.Contains(x))) return "CF";
    if (p.Contains("text") || p.Contains("title") || p.Contains("label")) return "Text";
    return "Value";
}

void RecordMeasure(string measureName, List<string> path, string reportName)
{
    Measure ms;
    if (measureName == null || !measuresByName.TryGetValue(measureName, out ms)) return;
    var sig = signals[ms];
    sig.References++;
    int projections = path.IndexOf("projections");
    int formattingRoot = Math.Max(path.IndexOf("objects"), path.IndexOf("visualContainerObjects"));
    if (projections > 0)
    {
        var role = path[projections - 1];
        if (role == "Tooltips") { sig.Tooltip = true; sig.Details.Add("Tooltip"); }
        else { sig.Projection = true; sig.Details.Add("Visual (" + role + ")"); }
    }
    else if (path.Contains("sortDefinition")) { sig.Sort = true; sig.Details.Add("Sort"); }
    else if (path.Contains("filterConfig")) { sig.Filter = true; sig.Details.Add("Filter"); }
    else if (formattingRoot >= 0)
    {
        var objectName = formattingRoot + 1 < path.Count ? path[formattingRoot + 1] : "";
        var propsIndex = path.IndexOf("properties", formattingRoot);
        var propertyName = propsIndex >= 0 && propsIndex + 1 < path.Count ? path[propsIndex + 1] : "";
        var kind = FormattingKind(propertyName);
        if (kind == "Text") { sig.Text = true; sig.Details.Add("Dynamic text (" + objectName + "." + propertyName + ")"); }
        else if (kind == "CF") { sig.ConditionalFormatting = true; sig.Details.Add("Conditional formatting (" + objectName + "." + propertyName + ")"); }
        else { sig.Projection = true; sig.Details.Add("Shown in visual (" + objectName + "." + propertyName + ")"); }
    }
    else { sig.Projection = true; sig.Details.Add("Visual"); }
}

void Walk(JToken token, List<string> path, Dictionary<string, string> aliases, string reportName)
{
    var obj = token as JObject;
    if (obj != null)
    {
        var localAliases = aliases;
        var fromArray = obj["From"] as JArray;
        if (fromArray != null)
        {
            localAliases = new Dictionary<string, string>(aliases, StringComparer.OrdinalIgnoreCase);
            foreach (var from in fromArray.OfType<JObject>())
                if ((string)from["Name"] != null && (string)from["Entity"] != null) localAliases[(string)from["Name"]] = (string)from["Entity"];
        }
        var parentKey = path.Count > 0 ? path[path.Count - 1] : null;
        if (parentKey == "Measure" || parentKey == "Column")
        {
            var table = (string)obj["Expression"]?["SourceRef"]?["Entity"];
            var alias = (string)obj["Expression"]?["SourceRef"]?["Source"];
            if (table == null && alias != null) localAliases.TryGetValue(alias, out table);
            var field = (string)obj["Property"];
            if (parentKey == "Measure") RecordMeasure(field, path, reportName);
            else if (table != null && field != null) usedColumns.Add(table + "|" + field);
        }
        foreach (var property in obj.Properties())
        {
            path.Add(property.Name);
            Walk(property.Value, path, localAliases, reportName);
            path.RemoveAt(path.Count - 1);
        }
    }
    var array = token as JArray;
    if (array != null) foreach (var item in array) Walk(item, path, aliases, reportName);
}

foreach (var report in selectedReports)
{
    var definitionFolder = Path.Combine(report.ReportFolder, "definition");
    if (!Directory.Exists(definitionFolder)) { Warning("Skipping report without definition folder: " + report.ReportFolder); continue; }
    foreach (var file in Directory.GetFiles(definitionFolder, "*.json", SearchOption.AllDirectories)
        .Where(fp => fp.EndsWith("report.json", StringComparison.OrdinalIgnoreCase) || fp.EndsWith("page.json", StringComparison.OrdinalIgnoreCase) || fp.EndsWith("visual.json", StringComparison.OrdinalIgnoreCase)))
    {
        try
        {
            Walk(JObject.Parse(File.ReadAllText(file)), new List<string>(), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), report.ReportName);
            scannedFiles++;
        }
        catch (Exception ex) { skippedFiles++; Warning("Skipping unreadable PBIR JSON file:\n" + file + "\n\n" + ex.Message); }
    }
}

// Fabric Apps: a measure an app queries is shown to its users, like a value in a
// visual (KPI). .dax files are parsed for Table[Column], 'Table'[Column],
// Table[Measure] and [Measure]; .ts/.tsx files only add names that exist in the
// model. Query result aliases and query measures (DEFINE MEASURE) are ignored.
var appNotInModel = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
int scannedAppFiles = 0;
var StringLiteral = new Regex("\"((?:[^\"]|\"\")*)\"");
var QualifiedRef = new Regex(@"(?<![\w\].])('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]");
var BareRef = new Regex(@"\[((?:[^\]]|\]\])+)\]");
var DefineMeasure = new Regex(@"\bMEASURE\s+('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]\s*=", RegexOptions.IgnoreCase);
string UnquoteTable(string t) { return t.StartsWith("'") ? t.Substring(1, t.Length - 2).Replace("''", "'") : t; }
string UnescapeName(string n) { return n.Replace("]]", "]"); }

// Removes DAX comments and string literals; keeps 'quoted tables' and [names] intact.
string StripDaxCommentsAndStrings(string dax)
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
}

var appTablesByName = Model.Tables.ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);
Column FindAppColumn(string table, string name)
{
    Table t;
    if (!appTablesByName.TryGetValue(table, out t)) return null;
    return t.Columns.FirstOrDefault(c => SameName(c.Name, name));
}
void RecordAppMeasure(Measure ms, string appName)
{
    var sig = signals[ms];
    sig.App = true;
    sig.References++;
    sig.Details.Add("Fabric App (" + appName + ")");
}
// Columns only matter here for field parameters and calculation groups used by an app
void RecordAppColumn(Column col) { usedColumns.Add(col.Table.Name + "|" + col.Name); }

foreach (var app in selectedApps)
{
    var appFiles = EnumerateSourceFiles(app.ReportFolder);

    foreach (var file in appFiles.Where(p => p.EndsWith(".dax", StringComparison.OrdinalIgnoreCase)))
    {
        scannedAppFiles++;
        var rel = app.ReportName + "/" + file.Substring(app.ReportFolder.Length).TrimStart('\\', '/').Replace('\\', '/');
        var rawDax = File.ReadAllText(file);
        var dax = StripDaxCommentsAndStrings(rawDax);

        // Result-column names created in the query ("Week" in SELECTCOLUMNS / ROW / ...)
        var queryAliases = new HashSet<string>(
            StringLiteral.Matches(rawDax).Cast<Match>().Select(x => x.Groups[1].Value.Replace("\"\"", "\"")),
            StringComparer.OrdinalIgnoreCase);
        var queryMeasures = new HashSet<string>(
            DefineMeasure.Matches(dax).Cast<Match>().Select(x => UnescapeName(x.Groups[2].Value)),
            StringComparer.OrdinalIgnoreCase);

        var remaining = new System.Text.StringBuilder(dax);
        foreach (Match m in QualifiedRef.Matches(dax))
        {
            var table = UnquoteTable(m.Groups[1].Value);
            if (!appTablesByName.ContainsKey(table)) continue; // keyword (ORDER BY [x]) or VAR: bare pass
            var name = UnescapeName(m.Groups[2].Value);
            for (int k = m.Index; k < m.Index + m.Length; k++) remaining[k] = ' ';
            if (queryMeasures.Contains(name)) continue;

            var col = FindAppColumn(table, name);
            Measure meas;
            if (col != null) RecordAppColumn(col);
            else if (measuresByName.TryGetValue(name, out meas)) RecordAppMeasure(meas, app.ReportName);
            else appNotInModel.Add(table + "[" + name + "]  (" + rel + ")");
        }

        foreach (Match m in BareRef.Matches(remaining.ToString()))
        {
            var name = UnescapeName(m.Groups[1].Value);
            Measure meas;
            if (measuresByName.TryGetValue(name, out meas)) RecordAppMeasure(meas, app.ReportName);
            else if (!queryMeasures.Contains(name) && !queryAliases.Contains(name))
                appNotInModel.Add("[" + name + "]  (" + rel + ")");
        }
    }

    // .ts/.tsx: only names that exist in the model (DAX fragments built in code,
    // measure names handed to DAX templates). Tests and type declarations skipped.
    foreach (var file in appFiles.Where(p =>
        (p.EndsWith(".ts", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".tsx", StringComparison.OrdinalIgnoreCase))
        && !Regex.IsMatch(Path.GetFileName(p), @"\.(spec|test)\.tsx?$", RegexOptions.IgnoreCase)
        && !p.EndsWith(".d.ts", StringComparison.OrdinalIgnoreCase)))
    {
        scannedAppFiles++;
        var code = File.ReadAllText(file);
        foreach (Match m in BareRef.Matches(code))
        {
            Measure meas;
            if (measuresByName.TryGetValue(UnescapeName(m.Groups[1].Value), out meas)) RecordAppMeasure(meas, app.ReportName);
        }
        foreach (Match m in QualifiedRef.Matches(code))
        {
            var col = FindAppColumn(UnquoteTable(m.Groups[1].Value), UnescapeName(m.Groups[2].Value));
            if (col != null) RecordAppColumn(col);
        }
    }
}

// Field parameters used in visuals: the measures they list are shown in visuals too
foreach (var calcTable in Model.Tables.OfType<CalculatedTable>())
{
    if (!calcTable.Columns.Any(cl => usedColumns.Contains(calcTable.Name + "|" + cl.Name))) continue;
    foreach (var ms in DependenciesOf(calcTable).OfType<Measure>())
    {
        signals[ms].FieldParameter = true;
        signals[ms].Details.Add("Field parameter (" + calcTable.Name + ")");
    }
}

// Follows measure and DAX UDF references from the start objects and calls
// mark(measure, "used by ...") for every measure reached (directly or indirectly)
var objectLabels = new Dictionary<object, string>();
void MarkReachable(IEnumerable<object> start, Action<Measure, string> mark)
{
    var queue = new Queue<object>(start);
    var visited = new HashSet<object>(queue);
    while (queue.Count > 0)
    {
        var current = queue.Dequeue();
        foreach (var dep in DependenciesOf(current))
        {
            if (!(dep is Measure) && dep.GetType().Name != "Function") continue;
            if (!visited.Add(dep)) continue;
            var depMeasure = dep as Measure;
            if (depMeasure != null && signals.ContainsKey(depMeasure))
            {
                var currentMeasure = current as Measure;
                string label;
                mark(depMeasure, currentMeasure != null ? "[" + currentMeasure.Name + "]"
                    : objectLabels.TryGetValue(current, out label) ? label : "a DAX UDF");
            }
            queue.Enqueue(dep);
        }
    }
}

// Helper measures: referenced (directly or indirectly, also through DAX UDFs) by a used
// measure, or by a calculation item of a calculation group used in the reports
{
    var start = signals.Where(kv => HasVisualUsage(kv.Value)).Select(kv => (object)kv.Key).ToList();
    foreach (var calcGroup in Model.Tables.OfType<CalculationGroupTable>())
    {
        if (!calcGroup.Columns.Any(cl => usedColumns.Contains(calcGroup.Name + "|" + cl.Name))) continue;
        foreach (var item in calcGroup.CalculationItems)
        {
            objectLabels[item] = "calculation item '" + calcGroup.Name + "'[" + item.Name + "]";
            start.Add(item);
        }
    }
    MarkReachable(start, (ms, by) => { signals[ms].Helper = true; signals[ms].Details.Add("Used by " + by); });
}

// Perspectives: a measure in a perspective - or used by one - is a model dependency
// (downstream use such as perspective models), even without visual usage
{
    var perspectiveNames = Model.Perspectives.Select(p => p.Name).ToList();
    var start = new List<object>();
    foreach (var ms in Model.AllMeasures)
    {
        var names = perspectiveNames.Where(p => ms.InPerspective[p]).ToList();
        if (names.Count == 0) continue;
        signals[ms].Perspective = true;
        signals[ms].Details.Add("Perspective: " + string.Join(", ", names));
        start.Add(ms);
    }
    MarkReachable(start, (ms, by) => { signals[ms].Perspective = true; signals[ms].Details.Add("Used by " + by + " (in a perspective)"); });
}

// Duplicates: same DAX (ignoring whitespace, comments and case) as another measure
var duplicateOf = new Dictionary<Measure, string>();
foreach (var group in Model.AllMeasures.Where(ms => !string.IsNullOrWhiteSpace(ms.Expression)).GroupBy(ms => NormalizeDax(ms.Expression)).Where(g => g.Count() > 1))
    foreach (var ms in group)
        duplicateOf[ms] = string.Join(", ", group.Where(o => o != ms).Select(o => "[" + o.Name + "]").Take(3));

// ---------------------------------------------------------------------------
// Categories and classification
// ---------------------------------------------------------------------------
var categories = new List<Category> {
    new Category { Key = "CF",     Name = "Conditional formatting", Folder = "Conditional formatting", Contains = "", DetectedBy = "Used in visuals only for formatting (colors, icons, data bars, ...)" },
    new Category { Key = "Text",   Name = "Dynamic text",           Folder = "Dynamic text",           Contains = "", DetectedBy = "Used in visuals only as text (titles, subtitles, labels, button/text values)" },
    new Category { Key = "SVG",    Name = "SVG / images",           Folder = "SVG and images",         Contains = "", DetectedBy = "Data category ImageUrl, actually used in a visual" },
    new Category { Key = "Filter", Name = "Visual filter",          Folder = "Visual filters",         Contains = "",    DetectedBy = "Only used as a visual-level filter" },
    new Category { Key = "KPI",    Name = "KPI (used in visuals)",  Folder = "KPI",                    Contains = "",    DetectedBy = "In a visual well, sort, tooltip, field parameter, shown as a value (card reference label, error bar, ...), or queried by a Fabric App" },
    new Category { Key = "Helper", Name = "Helper measures",        Folder = "Helper measures",        Contains = "",    DetectedBy = "Referenced by a used measure, DAX UDF, or calculation item of a calculation group used in visuals" },
    new Category { Key = "Perspective", Name = "In perspective (model dependency)", Folder = "In perspective", Contains = "", DetectedBy = "In a perspective, or used by a measure in a perspective - no visual usage found" },
    new Category { Key = "None",   Name = "No visual dependency",   Folder = "No visual dependency",   Contains = "",    DetectedBy = "Everything else (always last)" }
};

bool NameMatches(Measure ms, Category cat)
{
    return (cat.Contains ?? "").Split(';').Select(t => t.Trim()).Where(t => t.Length > 0)
        .Any(t => ms.Name.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);
}

string DataCategoryOf(Measure ms)
{
    var prop = ms.GetType().GetProperty("DataCategory");
    return prop == null ? null : prop.GetValue(ms, null) as string;
}

// Returns the reason when the measure belongs to the category, otherwise null
string MatchReason(Measure ms, Category cat)
{
    var sig = signals[ms];
    var byName = NameMatches(ms, cat) ? "Name contains '" + cat.Contains + "'" : null;
    // Shown as a value somewhere (well, tooltip, sort, field parameter, card reference label, ...):
    // such a measure is a KPI even if it is also used for formatting or text elsewhere
    // A Fabric App query counts as shown as a value (apps have no formatting / text roles to tell apart)
    bool shownAsValue = sig.Projection || sig.Tooltip || sig.Sort || sig.FieldParameter || sig.App;
    switch (cat.Key)
    {
        case "CF": return sig.ConditionalFormatting && !shownAsValue ? "Used in visuals only for conditional formatting" : byName;
        case "Text": return sig.Text && !shownAsValue ? "Used in visuals only as dynamic text" : byName;
        case "SVG": return SameName(DataCategoryOf(ms), "ImageUrl") && HasVisualUsage(sig) ? "Data category ImageUrl, used in visuals" : byName;
        case "Filter": return sig.Filter && !shownAsValue ? "Only used as visual filter" : byName;
        case "KPI": return shownAsValue
            ? (sig.App && !(sig.Projection || sig.Tooltip || sig.Sort || sig.FieldParameter) ? "Used in Fabric Apps" : sig.App ? "Used in visuals and Fabric Apps" : "Used in visuals")
            : byName;
        case "Helper": return sig.Helper ? "Referenced by used measures or calculation items" : byName;
        case "Perspective": return sig.Perspective ? "Model dependency through a perspective" : byName;
        case "None": return "No usage found in the selected reports and apps";
        default: return byName;
    }
}

Category Classify(Measure ms, List<Category> ordered, out string reason)
{
    foreach (var cat in ordered.Where(c => c.Key != "None").Concat(ordered.Where(c => c.Key == "None")))
    {
        reason = MatchReason(ms, cat);
        if (reason != null) return cat;
    }
    reason = null;
    return null;
}

// ---------------------------------------------------------------------------
// Step 3: settings dialog
// ---------------------------------------------------------------------------
var selectedMeasures = Selected.Measures.ToList();
bool selectedOnly = selectedMeasures.Count > 0;
int placementMode = 0;                  // 0 = root + keep existing, 1 = root only, 2 = subfolder of existing
bool useDuplicateFolder = true;
string duplicateFolderName = "Duplicates";
bool writeAnnotation = true;
bool hideHelpers = false;

using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Organize measures by usage", 7, DialogSize(0.8, 0.85));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var scopeCheck = new CheckBox {
        Text = "Only the " + selectedMeasures.Count + " measure(s) selected in Tabular Editor (otherwise all " + Model.AllMeasures.Count() + " measures)",
        Checked = selectedOnly, Enabled = selectedMeasures.Count > 0, AutoSize = true, Margin = new Padding(0, 0, 0, 8)
    };

    // Categories grid + side buttons
    var gridRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
    gridRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    gridRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    var grid = new DataGridView {
        Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect, MultiSelect = false,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false, EditMode = DataGridViewEditMode.EditOnEnter
    };
    grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
    grid.DefaultCellStyle.ForeColor = Color.FromArgb(90, 90, 90);
    var applyColumn = new DataGridViewCheckBoxColumn { Name = "Apply", HeaderText = "Apply", Width = 55 };
    var nameColumn = new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "Category (edit)", Width = 190 };
    var folderColumn = new DataGridViewTextBoxColumn { Name = "Folder", HeaderText = "Folder name (edit)", Width = 200 };
    var containsColumn = new DataGridViewTextBoxColumn { Name = "Contains", HeaderText = "Name contains (edit, ; separated)", Width = 200 };
    foreach (var col in new DataGridViewColumn[] { applyColumn, nameColumn, folderColumn, containsColumn }) Highlight(col);
    grid.Columns.AddRange(new DataGridViewColumn[] { applyColumn, nameColumn, folderColumn, containsColumn });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Count", HeaderText = "Measures", Width = 75, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Detected", HeaderText = "Detected by (in this order, first match wins)", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200 });
    grid.CurrentCellDirtyStateChanged += (s, e) => { if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };

    var sideButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(8, 0, 0, 0) };
    var upButton = new Button { Text = "Move up", Width = 110 };
    var downButton = new Button { Text = "Move down", Width = 110 };
    var addButton = new Button { Text = "Add category", Width = 110, Margin = new Padding(3, 16, 3, 3) };
    var removeButton = new Button { Text = "Remove category", Width = 110 };
    sideButtons.Controls.AddRange(new Control[] { upButton, downButton, addButton, removeButton });
    gridRow.Controls.Add(grid, 0, 0);
    gridRow.Controls.Add(sideButtons, 1, 0);

    // Placement and options
    var placementBox = new GroupBox { Text = "Folder placement", AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 0) };
    var placementFlow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Dock = DockStyle.Fill };
    var modeRootKeep = new RadioButton { Text = "Category as root folder, keep existing folders below it  (KPI\\Sales\\Margin)", AutoSize = true, Checked = true };
    var modeRootOnly = new RadioButton { Text = "Category as root folder only  (KPI) - the original folder is kept in an annotation", AutoSize = true };
    var modeSuffix = new RadioButton { Text = "Category as subfolder of the existing folder  (Sales\\Margin\\KPI)", AutoSize = true };
    placementFlow.Controls.AddRange(new Control[] { modeRootKeep, modeRootOnly, modeSuffix });
    placementBox.Controls.Add(placementFlow);

    var optionsFlow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, Margin = new Padding(0, 8, 0, 0) };
    var duplicateCheck = new CheckBox { Text = "Duplicates (same DAX as another measure) under No visual dependency in subfolder:", AutoSize = true, Checked = true, Margin = new Padding(0, 4, 4, 0) };
    var duplicateBox = new TextBox { Text = duplicateFolderName, Width = 160, BackColor = editBack };
    var annotationCheck = new CheckBox { Text = "Write annotation " + AnnUsage + " = <category> (for model documentation)", AutoSize = true, Checked = true, Margin = new Padding(24, 4, 0, 0) };
    var hideHelperCheck = new CheckBox { Text = "Hide the measures in Helper measures (only needed by other measures)", AutoSize = true, Checked = false, Margin = new Padding(0, 8, 0, 0) };
    optionsFlow.Controls.AddRange(new Control[] { duplicateCheck, duplicateBox, annotationCheck });
    optionsFlow.SetFlowBreak(annotationCheck, true);
    optionsFlow.Controls.Add(hideHelperCheck);

    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
    var previewButton = new Button { Text = "Preview", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { previewButton, cancelButton });

    bool loading = false;
    Action fillGrid = () =>
    {
        loading = true;
        grid.Rows.Clear();
        foreach (var cat in categories)
        {
            int index = grid.Rows.Add(cat.Apply, cat.Name, cat.Folder, cat.Contains, "", cat.DetectedBy ?? "Name contains only");
            grid.Rows[index].Tag = cat;
            if (cat.Key == "None")
            {
                var cell = grid.Rows[index].Cells["Contains"];
                cell.ReadOnly = true; cell.Style.BackColor = naBack; cell.Style.SelectionBackColor = naBack;
            }
        }
        loading = false;
    };
    Action refreshCounts = () =>
    {
        var inScope = scopeCheck.Checked ? selectedMeasures : Model.AllMeasures.ToList();
        var counts = new Dictionary<Category, int>();
        foreach (var ms in inScope)
        {
            string reason;
            var cat = Classify(ms, categories, out reason);
            if (cat != null) { int n; counts.TryGetValue(cat, out n); counts[cat] = n + 1; }
        }
        foreach (DataGridViewRow row in grid.Rows)
        {
            int n;
            counts.TryGetValue((Category)row.Tag, out n);
            row.Cells["Count"].Value = n;
        }
    };
    grid.CellValueChanged += (s, e) =>
    {
        if (loading || e.RowIndex < 0) return;
        var cat = (Category)grid.Rows[e.RowIndex].Tag;
        var row = grid.Rows[e.RowIndex];
        cat.Apply = Convert.ToBoolean(row.Cells["Apply"].Value);
        cat.Name = Convert.ToString(row.Cells["Category"].Value).Trim();
        cat.Folder = Convert.ToString(row.Cells["Folder"].Value).Trim().Replace("/", "\\").Trim('\\');
        cat.Contains = Convert.ToString(row.Cells["Contains"].Value);
        refreshCounts();
    };
    Action<int> move = delta =>
    {
        if (grid.CurrentCell == null) return;
        int index = grid.CurrentCell.RowIndex;
        var cat = (Category)grid.Rows[index].Tag;
        int target = index + delta;
        if (cat.Key == "None" || target < 0 || target >= categories.Count || categories[target].Key == "None") return;
        categories.RemoveAt(index);
        categories.Insert(target, cat);
        fillGrid();
        refreshCounts();
        grid.CurrentCell = grid.Rows[target].Cells["Category"];
    };
    upButton.Click += (s, e) => move(-1);
    downButton.Click += (s, e) => move(1);
    addButton.Click += (s, e) =>
    {
        // New name-based category, placed just before "No visual dependency"
        var cat = new Category { Key = "Custom", Name = "New category", Folder = "New category", Contains = "", DetectedBy = "Name contains only" };
        categories.Insert(categories.FindIndex(c => c.Key == "None"), cat);
        fillGrid();
        refreshCounts();
        var index = categories.IndexOf(cat);
        grid.CurrentCell = grid.Rows[index].Cells["Category"];
        grid.BeginEdit(true);
    };
    removeButton.Click += (s, e) =>
    {
        if (grid.CurrentCell == null) return;
        var cat = (Category)grid.Rows[grid.CurrentCell.RowIndex].Tag;
        if (cat.Key != "Custom") { MessageBox.Show("Built-in categories cannot be removed - untick Apply to leave their measures where they are.", "Remove category", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        categories.Remove(cat);
        fillGrid();
        refreshCounts();
    };
    scopeCheck.CheckedChanged += (s, e) => refreshCounts();
    duplicateCheck.CheckedChanged += (s, e) => duplicateBox.Enabled = duplicateCheck.Checked;

    previewButton.Click += (s, e) =>
    {
        grid.EndEdit();
        foreach (var cat in categories.Where(c => c.Apply))
        {
            if (string.IsNullOrWhiteSpace(cat.Folder)) { MessageBox.Show("Enter a folder name for '" + cat.Name + "', or untick Apply.", "Folder name", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (cat.Key == "Custom" && string.IsNullOrWhiteSpace(cat.Contains)) { MessageBox.Show("Enter the 'name contains' text for '" + cat.Name + "'.", "Name contains", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        }
        if (duplicateCheck.Checked && string.IsNullOrWhiteSpace(duplicateBox.Text)) { MessageBox.Show("Enter a folder name for duplicates, or untick the option.", "Duplicates", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        selectedOnly = scopeCheck.Checked;
        placementMode = modeRootKeep.Checked ? 0 : modeRootOnly.Checked ? 1 : 2;
        useDuplicateFolder = duplicateCheck.Checked;
        duplicateFolderName = duplicateBox.Text.Trim().Replace("/", "\\").Trim('\\');
        writeAnnotation = annotationCheck.Checked;
        hideHelpers = hideHelperCheck.Checked;
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(Heading("Organize measures into folders by how they are used in the " + selectedReports.Count + " selected report(s) (" + scannedFiles + " PBIR files scanned" + (skippedFiles > 0 ? ", " + skippedFiles + " skipped" : "") + ")."), 0, 0);
    panel.Controls.Add(scopeCheck, 0, 1);
    panel.Controls.Add(gridRow, 0, 2);
    panel.Controls.Add(placementBox, 0, 3);
    panel.Controls.Add(optionsFlow, 0, 4);
    panel.Controls.Add(new Label {
        Text = "Categories are checked from top to bottom - a measure goes to the first category it matches. 'Name contains' also puts a measure in the category. " +
               "Untick Apply to leave that category's measures in their current folder. Only the reports and apps in the selected folders are analysed - usage elsewhere (Service, Copilot, Excel, notebooks) is not detected.",
        AutoSize = true, MaximumSize = new Size(1200, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 0)
    }, 0, 5);
    panel.Controls.Add(buttonRow, 0, 6);
    dlg.CancelButton = cancelButton;
    fillGrid();
    refreshCounts();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 4: compute new folders
// ---------------------------------------------------------------------------
// Folder names to strip before placing a measure: current and earlier category folders, the folders of older macro 16 / 22 versions
var knownFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "XX - Not to be used", "Not used in report" };
foreach (var cat in categories) if (!string.IsNullOrWhiteSpace(cat.Folder)) foreach (var part in cat.Folder.Split('\\')) knownFolders.Add(part.Trim());
if (!string.IsNullOrWhiteSpace(duplicateFolderName)) knownFolders.Add(duplicateFolderName);
var previousFolders = Model.GetAnnotation(AnnKnownFolders);
if (!string.IsNullOrWhiteSpace(previousFolders)) foreach (var part in previousFolders.Split('|')) if (part.Trim().Length > 0) knownFolders.Add(part.Trim());

// The measure's own folder, without category folders:
//   1. still in the folder this macro set last time -> the folder remembered then
//   2. otherwise category folder names (current and earlier) and old macro 16 / 22 folders
//      are removed from the start and end of the path - a folder of your own with
//      the same name deeper in the path is kept
//   3. nothing left -> the folder remembered by this macro or by old macro 16 / 22
string BaseFolder(Measure ms)
{
    var current = ms.DisplayFolder ?? "";
    var applied = ms.GetAnnotation(AnnAppliedFolder);
    if (applied != null && string.Equals(applied, current, StringComparison.OrdinalIgnoreCase))
        return ms.GetAnnotation(AnnOriginalFolder) ?? "";

    var parts = current.Split('\\').Select(p => p.Trim()).Where(p => p.Length > 0).ToList();
    while (parts.Count > 0 && knownFolders.Contains(parts[0])) parts.RemoveAt(0);
    while (parts.Count > 0 && knownFolders.Contains(parts[parts.Count - 1])) parts.RemoveAt(parts.Count - 1);
    var basePath = string.Join("\\", parts);
    if (basePath.Length == 0) basePath = ms.GetAnnotation(AnnOriginalFolder) ?? ms.GetAnnotation(AnnMacro16OriginalFolder) ?? "";
    return basePath;
}

string Combine(params string[] parts) { return string.Join("\\", parts.Where(p => !string.IsNullOrWhiteSpace(p))); }

var inScopeMeasures = selectedOnly ? selectedMeasures : Model.AllMeasures.ToList();
var plans = new List<MeasurePlan>();
foreach (var ms in inScopeMeasures)
{
    string reason;
    var cat = Classify(ms, categories, out reason);
    var plan = new MeasurePlan { Measure = ms, Category = cat, Reason = reason };
    string dup;
    if (cat != null && cat.Key == "None" && useDuplicateFolder && duplicateOf.TryGetValue(ms, out dup)) { plan.Duplicate = true; plan.DuplicateOf = dup; }
    var sig = signals[ms];
    if (sig.Details.Count > 0) plan.Reason += "  |  " + string.Join("; ", sig.Details.Take(4)) + (sig.Details.Count > 4 ? "; ..." : "");

    if (cat == null || !cat.Apply) plan.NewFolder = ms.DisplayFolder ?? "";
    else
    {
        var categoryFolder = Combine(cat.Folder, plan.Duplicate ? duplicateFolderName : null);
        var basePath = BaseFolder(ms);
        plan.BaseFolder = basePath;
        plan.NewFolder = placementMode == 0 ? Combine(categoryFolder, basePath)
            : placementMode == 1 ? categoryFolder
            : Combine(basePath, categoryFolder);
    }
    plan.Hide = hideHelpers && cat != null && cat.Key == "Helper" && !ms.IsHidden;
    plans.Add(plan);
}

// ---------------------------------------------------------------------------
// Step 5: preview with filter box
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
foreach (var column in new[] { "Table", "Measure", "Category", "Current folder", "New folder", "Changed", "Hidden", "Duplicate of", "Why" }) preview.Columns.Add(column);
foreach (var plan in plans.OrderBy(p => categories.IndexOf(p.Category)).ThenBy(p => p.Measure.Table.Name).ThenBy(p => p.Measure.Name))
    preview.Rows.Add(plan.Measure.Table.Name, plan.Measure.Name,
        plan.Category == null ? "" : plan.Category.Name + (plan.Category.Apply ? "" : " (not applied)"),
        plan.Measure.DisplayFolder ?? "", plan.NewFolder,
        string.Equals(plan.NewFolder, plan.Measure.DisplayFolder ?? "", StringComparison.Ordinal) && !plan.Hide ? "No" : "Yes",
        plan.Hide ? "Will be hidden" : plan.Measure.IsHidden ? "Yes" : "No",
        plan.DuplicateOf ?? "", plan.Reason ?? "");

int changes = plans.Count(p => !string.Equals(p.NewFolder, p.Measure.DisplayFolder ?? "", StringComparison.Ordinal));
int hides = plans.Count(p => p.Hide);
using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Preview - organize measures", 4, DialogSize(0.9, 0.85));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var view = new System.Data.DataView(preview);
    var filterRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 6) };
    var filterBox = new TextBox { Width = 320 };
    var changedOnly = new CheckBox { Text = "Only changed", AutoSize = true, Margin = new Padding(16, 4, 0, 0) };
    filterRow.Controls.AddRange(new Control[] { new Label { Text = "Filter:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) }, filterBox, changedOnly });
    var previewGrid = new DataGridView {
        Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, BackgroundColor = SystemColors.Window, DataSource = view
    };
    Action applyFilter = () =>
    {
        var text = (filterBox.Text ?? "").Trim().Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]");
        var parts = new List<string>();
        if (text.Length > 0)
            parts.Add("(" + string.Join(" OR ", new[] { "Table", "Measure", "Category", "[Current folder]", "[New folder]", "Why" }
                .Select(c => (c.StartsWith("[") ? c : "[" + c + "]") + " LIKE '%" + text + "%'")) + ")");
        if (changedOnly.Checked) parts.Add("Changed = 'Yes'");
        view.RowFilter = string.Join(" AND ", parts);
    };
    filterBox.TextChanged += (s, e) => applyFilter();
    changedOnly.CheckedChanged += (s, e) => applyFilter();

    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var applyButton = new Button { Text = "Apply", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0), Enabled = changes > 0 || hides > 0 || writeAnnotation };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { applyButton, cancelButton });

    var summaryText = string.Join("   ", plans.Where(p => p.Category != null).GroupBy(p => p.Category).OrderBy(g => categories.IndexOf(g.Key)).Select(g => g.Key.Name + ": " + g.Count()));
    var sourceText = "Scanned: " + selectedReports.Count + " PBIR report(s) (" + scannedFiles + " JSON files" + (skippedFiles > 0 ? ", " + skippedFiles + " skipped" : "") + "), " +
        selectedApps.Count + " Fabric App(s) (" + scannedAppFiles + " files)" +
        (appNotInModel.Count > 0 ? ".  App references not found in the model: " + appNotInModel.Count + " (" + string.Join(", ", appNotInModel.Take(3)) + (appNotInModel.Count > 3 ? ", ..." : "") + ")" : "");
    panel.Controls.Add(Heading(plans.Count + " measure(s), " + changes + " move to another folder" + (hides > 0 ? ", " + hides + " will be hidden" : "") + ". Click a column header to sort.\n" + summaryText + "\n" + sourceText), 0, 0);
    panel.Controls.Add(filterRow, 0, 1);
    panel.Controls.Add(previewGrid, 0, 2);
    panel.Controls.Add(buttonRow, 0, 3);
    dlg.CancelButton = cancelButton;
    dlg.Shown += (s, e) => filterBox.Focus();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 6: apply (undoable with Ctrl+Z in Tabular Editor)
// ---------------------------------------------------------------------------
int moved = 0, hidden = 0;
foreach (var plan in plans)
{
    var ms = plan.Measure;
    if (plan.Category != null && plan.Category.Apply)
    {
        // Remember the measure's own folder and the folder set now, so a later run
        // (also with other category names or placement) starts from the right folder
        if (plan.BaseFolder.Length > 0) ms.SetAnnotation(AnnOriginalFolder, plan.BaseFolder);
        else ms.RemoveAnnotation(AnnOriginalFolder);
        ms.SetAnnotation(AnnAppliedFolder, plan.NewFolder);

        if (!string.Equals(plan.NewFolder, ms.DisplayFolder ?? "", StringComparison.Ordinal))
        {
            ms.DisplayFolder = plan.NewFolder;
            moved++;
        }
    }
    if (plan.Hide) { ms.IsHidden = true; hidden++; }
    if (writeAnnotation && plan.Category != null)
        ms.SetAnnotation(AnnUsage, plan.Category.Name + (plan.Duplicate ? " (duplicate)" : ""));
}

// Remember the folder names so a later run with other names can still clean them up
knownFolders.Remove("XX - Not to be used");
knownFolders.Remove("Not used in report");
Model.SetAnnotation(AnnKnownFolders, string.Join("|", knownFolders.OrderBy(x => x)));

Info("Done. Moved " + moved + " measure(s) to category folders" + (hidden > 0 ? ", hid " + hidden + " helper measure(s)" : "") +
    (writeAnnotation ? " and wrote the " + AnnUsage + " annotation on " + plans.Count(p => p.Category != null) + " measure(s)" : "") + ".\n\n" +
    "Undo with Ctrl+Z in Tabular Editor. Save the model to keep the changes.\n" +
    "Tip: in model documentation, filter on the " + AnnUsage + " annotation to show only measures in use.");
