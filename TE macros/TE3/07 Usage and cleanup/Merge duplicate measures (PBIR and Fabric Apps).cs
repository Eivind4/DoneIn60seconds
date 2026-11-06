// ============================================================================
// 23 - Merge duplicate measures (PBIR and Fabric Apps).cs
//
// Finds measures with the same DAX expression and merges them one group at a
// time: pick the measure to keep, repoint the reports (and optionally Fabric
// Apps) to it, and delete the other measures from the model.
//
// Input:
//   - Entire model, or the measures selected in Tabular Editor. With selected
//     measures, every duplicate group that contains a selected measure is
//     processed; its duplicates can be anywhere in the model.
//
// Duplicate detection (same idea as the "Duplicate / redundant" column in 22,
// but stricter): comments are ignored, whitespace and case are ignored outside
// string literals, string literals must match exactly ("MM" and "mm" differ in
// FORMAT). Format strings, folders and descriptions may differ - the group
// dialog shows the differences.
//
// Step by step, for each duplicate group:
//   - Keep: defaults to the measure with the most references
//     (PBIR + Fabric App + model references).
//   - Merge + delete: the other measures (untick to leave one alone).
//   - "Merge this group" or "Do nothing for this group". Back / Finish / Cancel
//     are available on every step. Nothing is changed until the final preview
//     is confirmed.
//
// What a merge changes:
//   - Model: DAX that references a removed measure (measures, calculated
//     columns/tables, calculation items, RLS filters, KPIs, DAX UDFs, field
//     parameters) is rewritten to [Kept]. Perspective membership of the removed
//     measures is added to the kept one, and its description is copied when the
//     kept measure has none. The removed measures are deleted (Ctrl+Z works).
//   - PBIR (every .json below each selected report's definition folder,
//     including bookmarks): Measure references (Entity / From alias +
//     Property), queryRef, nativeQueryRef when it equals the old name,
//     selector.metadata, DAX in reportExtensions.json and visual calculations.
//     projection.displayName is kept.
//   - Fabric Apps (EXPERIMENTAL - NOT THOROUGHLY TESTED): [Measure] and
//     'Table'[Measure] in .dax files (query measures defined with DEFINE
//     MEASURE are left alone) and in .ts/.tsx files. Review the line-by-line
//     preview carefully.
//
// Not handled:
//   - Measure names in strings, e.g. SELECTEDMEASURENAME() = "Old" in
//     calculation items, or measure names built in app code.
//   - Reports and apps outside the selected folders, PBIX / PBIR-Legacy
//     reports, and anything else downstream in the Service (see macro 22).
//   - A visual that used both the kept and a removed measure ends up with the
//     same field twice; the preview warns, fix it in Power BI Desktop.
//
// IMPORTANT:
//   - Close Power BI Desktop (or use Desktop Bridge afterwards) and commit to
//     git first. PBIR and app edits are NOT undoable in Tabular Editor.
//   - After running, save the model (and deploy it if reports are
//     live-connected), then reopen the reports.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

const string UnusedFolder = "XX - Not to be used";   // root folder used by macro 22

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

class DuplicateGroup
{
    public List<Measure> Members = new List<Measure>();
    public Measure Keep;
    public List<Measure> Remove = new List<Measure>();
    public string Decision;   // null (not visited), "Merge" or "Skip"
}

class PbirDocument
{
    public string ReportName;
    public string DefinitionFolder;
    public string File;
    public JObject Root;
    public bool IsExtensions;
    public int Changes;
    public string Warning;
}

class AppDocument
{
    public string AppName;
    public string AppFolder;
    public string File;
    public bool IsDax;
    public string Text;
    public string NewText;
}

class ModelEdit
{
    public object Target;
    public string Property;
    public string Label;
    public string Before;
    public string After;
}

Func<Size> BigDialogSize = () =>
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * 0.85), (int)(wa.Height * 0.85));
};

Func<string, string> CleanReportName = folder =>
{
    var folderName = Path.GetFileName(folder);
    return folderName != null && folderName.EndsWith(".Report", StringComparison.OrdinalIgnoreCase)
        ? folderName.Substring(0, folderName.Length - ".Report".Length)
        : folderName;
};

Func<string, string> ReadConnectionSummary = definitionFile =>
{
    try
    {
        var json = JObject.Parse(File.ReadAllText(definitionFile));
        var dataset = json["datasetReference"] ?? json["semanticModelReference"];
        var byPath = (string)dataset?["byPath"]?["path"];
        if (!string.IsNullOrWhiteSpace(byPath)) return "byPath: " + byPath;
        var conn = (string)dataset?["byConnection"]?["connectionString"];
        if (!string.IsNullOrWhiteSpace(conn)) return "byConnection/live";
    }
    catch { }
    return "connection unknown";
};

Func<string, string, bool> SameName = (a, b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

Func<object, IEnumerable<object>> GetReferencedBy = obj =>
{
    if (obj == null) return Enumerable.Empty<object>();
    var prop = obj.GetType().GetProperty("ReferencedBy");
    if (prop == null) return Enumerable.Empty<object>();
    var value = prop.GetValue(obj, null) as System.Collections.IEnumerable;
    return value == null ? Enumerable.Empty<object>() : value.Cast<object>();
};

// String property through reflection (not every object / TE version has every DAX property)
Func<object, string, string> GetStringProperty = (obj, propertyName) =>
{
    var prop = obj == null ? null : obj.GetType().GetProperty(propertyName);
    return prop != null && prop.PropertyType == typeof(string) ? prop.GetValue(obj, null) as string : null;
};

Func<string, int, string> Shorten = (text, max) =>
{
    var oneLine = Regex.Replace(text ?? "", @"\s+", " ").Trim();
    return oneLine.Length <= max ? oneLine : oneLine.Substring(0, max) + " …";
};

// Read-only grid with a heading and Continue/Cancel (Output() only offers Close).
Func<string, string, System.Data.DataTable, string, bool> ShowGrid = (title, heading, table, continueText) =>
{
    using (var form = new Form())
    {
        form.Text = title;
        form.Size = BigDialogSize();
        form.MinimumSize = new Size(760, 480);
        form.StartPosition = FormStartPosition.CenterScreen;
        var font = new Font("Segoe UI", 9);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var label = new Label {
            Text = heading, Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(1200, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
        };
        var grid = new DataGridView {
            Dock = DockStyle.Fill, Font = font, ReadOnly = true,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            BackgroundColor = SystemColors.Window, DataSource = table
        };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var ok = new Button { Text = continueText, AutoSize = true, Font = font, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
        var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { ok, cancel });
        layout.Controls.Add(label, 0, 0);
        layout.Controls.Add(grid, 0, 1);
        layout.Controls.Add(buttons, 0, 2);
        form.Controls.Add(layout);
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK;
    }
};

// ---------------------------------------------------------------------------
// DAX helpers
// ---------------------------------------------------------------------------
var measuresByName = new Dictionary<string, Measure>(StringComparer.OrdinalIgnoreCase);
foreach (var am in Model.AllMeasures)
    if (!measuresByName.ContainsKey(am.Name)) measuresByName.Add(am.Name, am);

Func<Measure, string> MeasureRef = msr => "[" + msr.Name.Replace("]", "]]") + "]";

// Is "[name]" qualified by "qualifierTable" a reference to measure msr?
// 'Home table'[Measure] is; 'Other table'[Name] is a column reference.
Func<Measure, string, bool> IsMeasureReference = (msr, qualifierTable) =>
    qualifierTable == null || (SameName(qualifierTable, msr.Table.Name) && !msr.Table.Columns.Contains(msr.Name));

// Walks DAX and calls replace(measure) for every reference to a model measure:
// [Measure] or 'Home table'[Measure]. When replace returns another measure, the
// reference (including its table qualifier) becomes [Other]. Comments and string
// literals are copied unchanged.
Func<string, Func<Measure, Measure>, string> RewriteDax = (dax, replace) =>
{
    if (string.IsNullOrEmpty(dax)) return dax;
    var sb = new System.Text.StringBuilder(dax.Length);
    int pos = 0;
    int qualifierStart = -1;
    string qualifier = null;   // table name directly in front of the next [name], if any
    while (pos < dax.Length)
    {
        char ch = dax[pos];
        char nx = pos + 1 < dax.Length ? dax[pos + 1] : '\0';
        if ((ch == '/' && nx == '/') || (ch == '-' && nx == '-'))
        {
            int lineEnd = dax.IndexOf('\n', pos);
            if (lineEnd < 0) lineEnd = dax.Length;
            sb.Append(dax, pos, lineEnd - pos);
            pos = lineEnd;
            qualifier = null;
            continue;
        }
        if (ch == '/' && nx == '*')
        {
            int blockEnd = dax.IndexOf("*/", pos + 2);
            blockEnd = blockEnd < 0 ? dax.Length : blockEnd + 2;
            sb.Append(dax, pos, blockEnd - pos);
            pos = blockEnd;
            qualifier = null;
            continue;
        }
        if (ch == '"' || ch == '\'' || ch == '[')
        {
            char close = ch == '[' ? ']' : ch;
            int tokenStart = pos++;
            while (pos < dax.Length)
            {
                if (dax[pos] == close && pos + 1 < dax.Length && dax[pos + 1] == close) { pos += 2; continue; }
                if (dax[pos] == close) { pos++; break; }
                pos++;
            }
            var token = dax.Substring(tokenStart, pos - tokenStart);
            var inner = token.Length >= 2 ? token.Substring(1, token.Length - 2) : "";
            if (ch == '\'')
            {
                qualifierStart = sb.Length;
                qualifier = inner.Replace("''", "'");
                sb.Append(token);
                continue;
            }
            if (ch == '[')
            {
                Measure found;
                if (measuresByName.TryGetValue(inner.Replace("]]", "]"), out found) && IsMeasureReference(found, qualifier))
                {
                    var target = replace(found);
                    if (target != null && target != found)
                    {
                        if (qualifier != null) sb.Length = qualifierStart;
                        sb.Append(MeasureRef(target));
                        qualifier = null;
                        continue;
                    }
                }
            }
            sb.Append(token);
            qualifier = null;
            continue;
        }
        if (char.IsLetter(ch) || ch == '_')
        {
            int wordStart = pos;
            while (pos < dax.Length && (char.IsLetterOrDigit(dax[pos]) || dax[pos] == '_')) pos++;
            var word = dax.Substring(wordStart, pos - wordStart);
            // Only an unquoted table name qualifies a [name]; keywords such as RETURN do not
            qualifier = Model.Tables.Contains(word) ? word : null;
            qualifierStart = sb.Length;
            sb.Append(word);
            continue;
        }
        qualifier = null;
        sb.Append(ch);
        pos++;
    }
    return sb.ToString();
};

// Same as RewriteDax for TypeScript: only [Measure] and 'Table'[Measure] / Table[Measure]
// patterns (DAX inside strings and template literals). Code such as row[x] is kept.
var TsRef = new Regex(@"('(?:[^'\r\n]|'')+'|\b[A-Za-z_]\w*)?\[((?:[^\]\r\n]|\]\])+)\]");
Func<string, Func<Measure, Measure>, string> RewriteTs = (code, replace) =>
    TsRef.Replace(code, match =>
    {
        var prefix = match.Groups[1].Value;
        var prefixTable = prefix.StartsWith("'") ? prefix.Substring(1, prefix.Length - 2).Replace("''", "'") : prefix;
        bool tableQualified = prefix.Length > 0 && Model.Tables.Contains(prefixTable);
        Measure found;
        if (!measuresByName.TryGetValue(match.Groups[2].Value.Replace("]]", "]"), out found)) return match.Value;
        if (tableQualified && !IsMeasureReference(found, prefixTable)) return match.Value;
        var target = replace(found);
        if (target == null || target == found) return match.Value;
        return (tableQualified ? "" : prefix) + MeasureRef(target);
    });

// Query measures (DEFINE MEASURE 'T'[Name] = ...) override model measures in that query
var DefineMeasure = new Regex(@"\bMEASURE\s+('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]\s*=", RegexOptions.IgnoreCase);

// DAX compared for duplicates: no comments, whitespace collapsed (and dropped next to
// operators and brackets), case ignored except inside string literals.
Func<string, string> NormalizeForCompare = dax =>
{
    var sb = new System.Text.StringBuilder();
    int pos = 0;
    bool pendingSpace = false;
    Func<char, bool> isWordChar = x => char.IsLetterOrDigit(x) || x == '_';
    while (pos < dax.Length)
    {
        char ch = dax[pos];
        char nx = pos + 1 < dax.Length ? dax[pos + 1] : '\0';
        if ((ch == '/' && nx == '/') || (ch == '-' && nx == '-'))
        {
            while (pos < dax.Length && dax[pos] != '\n') pos++;
            pendingSpace = true;
            continue;
        }
        if (ch == '/' && nx == '*')
        {
            int blockEnd = dax.IndexOf("*/", pos + 2);
            pos = blockEnd < 0 ? dax.Length : blockEnd + 2;
            pendingSpace = true;
            continue;
        }
        if (char.IsWhiteSpace(ch)) { pendingSpace = true; pos++; continue; }
        string token;
        if (ch == '"' || ch == '\'' || ch == '[')
        {
            char close = ch == '[' ? ']' : ch;
            int tokenStart = pos++;
            while (pos < dax.Length)
            {
                if (dax[pos] == close && pos + 1 < dax.Length && dax[pos + 1] == close) { pos += 2; continue; }
                if (dax[pos] == close) { pos++; break; }
                pos++;
            }
            token = dax.Substring(tokenStart, pos - tokenStart);
            if (ch != '"') token = token.ToUpperInvariant();
        }
        else
        {
            token = char.ToUpperInvariant(ch).ToString();
            pos++;
        }
        if (pendingSpace && sb.Length > 0 && isWordChar(sb[sb.Length - 1]) && isWordChar(token[0])) sb.Append(' ');
        pendingSpace = false;
        sb.Append(token);
    }
    return sb.ToString();
};

// ---------------------------------------------------------------------------
// Open model location and remembered folders (shared with macro 22)
// ---------------------------------------------------------------------------
string[] SkippedFolders = { "node_modules", "dist", ".agents", ".git", ".temp", ".vite", "coverage", "bin", "obj" };

Func<string, IEnumerable<string>> EnumerateSourceFiles = rootDir =>
{
    var result = new List<string>();
    var pending = new Stack<string>();
    pending.Push(rootDir);
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

string modelSourcePath = null;
try
{
    var handler = Model.GetType().GetProperty("Handler")?.GetValue(Model, null);
    modelSourcePath = handler?.GetType().GetProperty("Source")?.GetValue(handler, null) as string;
    if (modelSourcePath != null && !File.Exists(modelSourcePath) && !Directory.Exists(modelSourcePath)) modelSourcePath = null;
}
catch { }

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
    var db = Model.Database;
    var dbId = db == null ? null : db.GetType().GetProperty("ID")?.GetValue(db, null) as string;
    if (modelSourcePath == null && !string.IsNullOrEmpty(dbId)) modelKey = modelName + "|" + dbId;
}
catch { }

// Folders from macro 22's settings, read only
JObject savedSources = null;
try
{
    var settingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TabularEditor3", "MacroSettings", "dependency-scan-sources.json");
    if (File.Exists(settingsFile)) savedSources = JObject.Parse(File.ReadAllText(settingsFile))[modelKey] as JObject;
}
catch { }
if (savedSources == null) savedSources = new JObject();

string pbirFolder = (string)savedSources["pbirFolder"] ?? repoRoot ?? "";
string appsFolder = (string)savedSources["appsFolder"] ?? repoRoot ?? "";
bool onlyThisModelReports = ((bool?)savedSources["onlyThisModelReports"] ?? true) && modelFolder != null;
bool scanPbir = true;
bool scanApps = false;   // experimental, always opt-in

var selectedMeasureList = Selected.Measures.ToList();
bool useSelected = selectedMeasureList.Count > 0;

// ---------------------------------------------------------------------------
// Dialog: input and sources
// ---------------------------------------------------------------------------
using (var form = new Form())
{
    form.Text = "Merge duplicate measures";
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
    Action<Control, int> addSpan = (ctrl, span) => { layout.Controls.Add(ctrl); layout.SetColumnSpan(ctrl, span); };
    Func<string, Font, Color, Padding, Label> addText = (text, f, color, margin) =>
    {
        var lbl = new Label { Text = text, AutoSize = true, MaximumSize = new Size(720, 0), Font = f, ForeColor = color, Margin = margin };
        addSpan(lbl, 3);
        return lbl;
    };

    addText("Close Power BI Desktop (or use Desktop Bridge afterwards) and commit to git first. Report and app files changed by this macro cannot be undone in Tabular Editor.",
        boldFont, Color.DarkRed, new Padding(0, 0, 0, 10));

    addText("Measures to look at", boldFont, Color.Black, new Padding(0, 0, 0, 2));
    var rbSelected = new RadioButton { Text = "Duplicate groups containing the selected measures (" + selectedMeasureList.Count + " selected)", Checked = useSelected, Enabled = useSelected, AutoSize = true, Margin = new Padding(18, 2, 0, 0) };
    var rbModel = new RadioButton { Text = "Entire model", Checked = !useSelected, AutoSize = true, Margin = new Padding(18, 2, 0, 10) };
    addSpan(rbSelected, 3);
    addSpan(rbModel, 3);

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

    var chkPbir = new CheckBox { Text = "Update PBIR reports (definition.pbir)", Checked = scanPbir, AutoSize = true, Font = boldFont, Margin = new Padding(0, 6, 0, 0) };
    addSpan(chkPbir, 3);
    var pbirBox = folderRow("Folder", pbirFolder);
    var chkOnlyThisModel = new CheckBox {
        Text = modelFolder != null
            ? "Only reports whose definition.pbir points to this model (" + Path.GetFileName(modelFolder) + ")"
            : "Only reports pointing to this model (needs a model opened from a .SemanticModel folder)",
        Checked = onlyThisModelReports, AutoSize = true, Margin = new Padding(18, 4, 0, 6)
    };
    addSpan(chkOnlyThisModel, 3);

    var chkApps = new CheckBox { Text = "Update Fabric Apps (rayfin/rayfin.yml)", Checked = scanApps, AutoSize = true, Font = boldFont, Margin = new Padding(0, 10, 0, 0) };
    addSpan(chkApps, 3);
    var appsBox = folderRow("Folder", appsFolder);
    addText("EXPERIMENTAL - not thoroughly tested. Rewrites [Measure] references in .dax, .ts and .tsx files. Review every line in the preview.",
        boldFont, Color.DarkOrange, new Padding(18, 2, 0, 10));

    Action syncEnabled = () =>
    {
        pbirBox.Enabled = chkPbir.Checked;
        chkOnlyThisModel.Enabled = chkPbir.Checked && modelFolder != null;
        appsBox.Enabled = chkApps.Checked;
    };
    chkPbir.CheckedChanged += (s, e) => syncEnabled();
    chkApps.CheckedChanged += (s, e) => syncEnabled();
    syncEnabled();

    addText("Without reports or apps, only the model is merged - visuals that use a removed measure will break.",
        font, Color.DimGray, new Padding(0, 0, 0, 10));

    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { ok, cancel });
    addSpan(buttons, 3);
    ok.Click += (s, e) =>
    {
        string problem = null;
        if (chkPbir.Checked && !Directory.Exists(pbirBox.Text.Trim().Trim('"'))) problem = "The PBIR folder does not exist.";
        else if (chkApps.Checked && !Directory.Exists(appsBox.Text.Trim().Trim('"'))) problem = "The Fabric Apps folder does not exist.";
        else if (!chkPbir.Checked && !chkApps.Checked &&
            MessageBox.Show("No reports or apps selected. Only the model will be changed, and visuals using removed measures will break.\n\nContinue?",
                "Model only", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        if (problem != null) { MessageBox.Show(problem, "Check the inputs", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    form.Controls.Add(layout);
    form.AcceptButton = ok;
    form.CancelButton = cancel;
    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }

    useSelected = rbSelected.Checked;
    scanPbir = chkPbir.Checked;
    scanApps = chkApps.Checked;
    pbirFolder = pbirBox.Text.Trim().Trim('"');
    appsFolder = appsBox.Text.Trim().Trim('"');
    onlyThisModelReports = chkOnlyThisModel.Checked && modelFolder != null;
}

// ---------------------------------------------------------------------------
// Duplicate groups
// ---------------------------------------------------------------------------
var groups = Model.AllMeasures
    .Where(x => !string.IsNullOrWhiteSpace(x.Expression))
    .GroupBy(x => NormalizeForCompare(x.Expression))
    .Where(grp => grp.Count() > 1)
    .Where(grp => !useSelected || grp.Any(x => selectedMeasureList.Contains(x)))
    .Select(grp => new DuplicateGroup { Members = grp.OrderBy(x => x.Table.Name).ThenBy(x => x.Name).ToList() })
    .OrderBy(grp => grp.Members[0].Table.Name).ThenBy(grp => grp.Members[0].Name)
    .ToList();

if (groups.Count == 0)
{
    Info(useSelected
        ? "None of the selected measures has a duplicate (same DAX) in the model. No changes made."
        : "No measures with the same DAX were found. No changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Discover and select reports / apps
// ---------------------------------------------------------------------------
Func<string, bool> PointsToOpenModel = definitionFile =>
{
    if (modelFolder == null) return true;
    try
    {
        var json = JObject.Parse(File.ReadAllText(definitionFile));
        var dataset = json["datasetReference"] ?? json["semanticModelReference"];
        var byPath = (string)dataset?["byPath"]?["path"];
        if (string.IsNullOrWhiteSpace(byPath)) return false;
        var target = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(definitionFile), byPath)).TrimEnd('\\', '/');
        return string.Equals(target, modelFolder.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }
    catch { return false; }
};

var reportChoices = new List<ReportChoice>();
if (scanPbir)
    foreach (var f in EnumerateSourceFiles(pbirFolder).Where(p => string.Equals(Path.GetFileName(p), "definition.pbir", StringComparison.OrdinalIgnoreCase)))
    {
        if (onlyThisModelReports && !PointsToOpenModel(f)) continue;
        reportChoices.Add(new ReportChoice {
            Kind = "PBIR", DefinitionFile = f, ReportFolder = Path.GetDirectoryName(f),
            ReportName = CleanReportName(Path.GetDirectoryName(f)), ConnectionSummary = ReadConnectionSummary(f), Selected = true
        });
    }
if (scanApps)
    foreach (var f in EnumerateSourceFiles(appsFolder).Where(p =>
        string.Equals(Path.GetFileName(p), "rayfin.yml", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(Path.GetFileName(Path.GetDirectoryName(p)), "rayfin", StringComparison.OrdinalIgnoreCase)))
    {
        var appRoot = Path.GetDirectoryName(Path.GetDirectoryName(f));
        reportChoices.Add(new ReportChoice {
            Kind = "App", DefinitionFile = f, ReportFolder = appRoot,
            ReportName = Path.GetFileName(appRoot), ConnectionSummary = "experimental", Selected = true
        });
    }
reportChoices = reportChoices.OrderBy(r => r.Kind).ThenBy(r => r.ReportName).ThenBy(r => r.ReportFolder).ToList();

if ((scanPbir || scanApps) && reportChoices.Count == 0)
{
    if (MessageBox.Show("No reports or apps were found in the selected folders.\n\nContinue with the model only? Visuals using removed measures will break.",
            "Nothing found", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
    { Info("Cancelled - no changes made."); return; }
}
else if (reportChoices.Count > 0)
using (var form = new Form())
{
    form.Text = "Select reports and apps to update";
    form.Size = BigDialogSize();
    form.MinimumSize = new Size(760, 520);
    form.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, Padding = new Padding(12) };
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var note = new Label {
        Text = "Every report that uses a measure that will be removed must be updated, or its visuals break. [App] rows are experimental - not thoroughly tested.",
        Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(1100, 0), Font = font, ForeColor = Color.DarkSlateGray
    };
    var search = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 8, 0, 8) };
    var list = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = font, HorizontalScrollbar = true };
    var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var checkAll = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var clear = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Font = font, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { checkAll, clear, ok, cancel });

    bool suppress = false;
    Action rebuild = () =>
    {
        suppress = true;
        list.Items.Clear();
        var filter = (search.Text ?? "").Trim();
        foreach (var r in reportChoices.Where(r => filter == "" || r.Label.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0))
            list.Items.Add(r.Label, r.Selected);
        suppress = false;
    };
    list.ItemCheck += (s, e) =>
    {
        if (suppress) return;
        var label = (string)list.Items[e.Index];
        reportChoices.First(r => r.Label == label).Selected = e.NewValue == CheckState.Checked;
    };
    search.TextChanged += (s, e) => rebuild();
    checkAll.Click += (s, e) =>
    {
        foreach (var r in reportChoices.Where(r => string.IsNullOrEmpty(search.Text) || r.Label.IndexOf(search.Text, StringComparison.OrdinalIgnoreCase) >= 0))
            r.Selected = true;
        rebuild();
    };
    clear.Click += (s, e) => { foreach (var r in reportChoices) r.Selected = false; rebuild(); };

    layout.Controls.Add(note, 0, 0);
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
// Load PBIR and app files
// ---------------------------------------------------------------------------
var documents = new List<PbirDocument>();
foreach (var report in selectedReports)
{
    var definitionFolder = Path.Combine(report.ReportFolder, "definition");
    if (!Directory.Exists(definitionFolder)) { Warning("Skipping report without definition folder: " + report.ReportFolder); continue; }
    foreach (var f in Directory.GetFiles(definitionFolder, "*.json", SearchOption.AllDirectories))
    {
        try
        {
            documents.Add(new PbirDocument {
                ReportName = report.ReportName, DefinitionFolder = definitionFolder, File = f,
                Root = JObject.Parse(File.ReadAllText(f)),
                IsExtensions = string.Equals(Path.GetFileName(f), "reportExtensions.json", StringComparison.OrdinalIgnoreCase)
            });
        }
        catch (Exception ex) { Warning("Skipping unreadable PBIR JSON file (it is not updated):\n" + f + "\n\n" + ex.Message); }
    }
}

var appDocuments = new List<AppDocument>();
foreach (var app in selectedApps)
    foreach (var f in EnumerateSourceFiles(app.ReportFolder).Where(p =>
        p.EndsWith(".dax", StringComparison.OrdinalIgnoreCase) ||
        ((p.EndsWith(".ts", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".tsx", StringComparison.OrdinalIgnoreCase))
            && !Regex.IsMatch(Path.GetFileName(p), @"\.(spec|test)\.tsx?$", RegexOptions.IgnoreCase)
            && !p.EndsWith(".d.ts", StringComparison.OrdinalIgnoreCase))))
    {
        appDocuments.Add(new AppDocument {
            AppName = app.ReportName, AppFolder = app.ReportFolder, File = f,
            IsDax = f.EndsWith(".dax", StringComparison.OrdinalIgnoreCase), Text = File.ReadAllText(f)
        });
    }

// Rewrites one app file; measures redefined in a .dax query are left alone
Func<AppDocument, Func<Measure, Measure>, string> RewriteApp = (doc, replace) =>
{
    if (!doc.IsDax) return RewriteTs(doc.Text, replace);
    var queryMeasures = new HashSet<string>(
        DefineMeasure.Matches(doc.Text).Cast<Match>().Select(x => x.Groups[2].Value.Replace("]]", "]")),
        StringComparer.OrdinalIgnoreCase);
    return RewriteDax(doc.Text, x => queryMeasures.Contains(x.Name) ? null : replace(x));
};

// ---------------------------------------------------------------------------
// Reference counts (to pick the default measure to keep)
// ---------------------------------------------------------------------------
Func<string, string, Measure> FindMeasure = (table, measureName) =>
{
    Measure found;
    return table != null && measureName != null && measuresByName.TryGetValue(measureName, out found) && SameName(found.Table.Name, table) ? found : null;
};

var pbirRefs = new Dictionary<Measure, int>();
var appRefs = new Dictionary<Measure, int>();
Action<Dictionary<Measure, int>, Measure> addRef = (counts, msr) =>
{
    int n;
    counts.TryGetValue(msr, out n);
    counts[msr] = n + 1;
};
Func<Dictionary<Measure, int>, Measure, int> RefCount = (counts, msr) =>
{
    int n;
    return counts.TryGetValue(msr, out n) ? n : 0;
};

Action<JToken, string, Dictionary<string, string>, bool> countPbir = null;
countPbir = (token, parentKey, aliasMap, isExtensions) =>
{
    var obj = token as JObject;
    if (obj != null)
    {
        var aliases = aliasMap;
        var fromArray = obj["From"] as JArray;
        if (fromArray != null)
        {
            aliases = new Dictionary<string, string>(aliasMap, StringComparer.OrdinalIgnoreCase);
            foreach (var from in fromArray.OfType<JObject>())
                if ((string)from["Name"] != null && (string)from["Entity"] != null) aliases[(string)from["Name"]] = (string)from["Entity"];
        }
        if (parentKey == "Measure")
        {
            var sourceRef = obj["Expression"]?["SourceRef"];
            var table = (string)sourceRef?["Entity"];
            var alias = (string)sourceRef?["Source"];
            if (table == null && alias != null) aliases.TryGetValue(alias, out table);
            var found = FindMeasure(table, (string)obj["Property"]);
            if (found != null) addRef(pbirRefs, found);
        }
        foreach (var prop in obj.Properties())
        {
            if (isExtensions && prop.Name == "expression" && prop.Value.Type == JTokenType.String)
                RewriteDax((string)prop.Value, x => { addRef(pbirRefs, x); return null; });
            else
                countPbir(prop.Value, prop.Name, aliases, isExtensions);
        }
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array) countPbir(item, parentKey, aliasMap, isExtensions);
    }
};

foreach (var doc in documents)
    countPbir(doc.Root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), doc.IsExtensions);
foreach (var doc in appDocuments)
    RewriteApp(doc, x => { addRef(appRefs, x); return null; });

Func<Measure, int> ModelRefs = msr => GetReferencedBy(msr).Count();
Func<Measure, int> TotalRefs = msr => RefCount(pbirRefs, msr) + RefCount(appRefs, msr) + ModelRefs(msr);

foreach (var grp in groups)
{
    grp.Keep = grp.Members
        .OrderByDescending(TotalRefs)
        .ThenBy(x => x.IsHidden)
        .ThenBy(x => (x.DisplayFolder ?? "") == UnusedFolder)
        .ThenBy(x => x.Table.Name).ThenBy(x => x.Name)
        .First();
    grp.Remove = grp.Members.Where(x => x != grp.Keep).ToList();
}

// ---------------------------------------------------------------------------
// Step by step: one dialog per duplicate group
// ---------------------------------------------------------------------------
Func<Measure, List<Measure>, List<string>> GroupWarnings = (kept, removed) =>
{
    var warnings = new List<string>();
    foreach (var r in removed)
    {
        var label = "'" + r.Table.Name + "'[" + r.Name + "]";
        if ((r.FormatString ?? "") != (kept.FormatString ?? ""))
            warnings.Add(label + ": format string \"" + r.FormatString + "\" differs from the kept \"" + kept.FormatString + "\".");
        if ((GetStringProperty(r, "FormatStringExpression") ?? "") != (GetStringProperty(kept, "FormatStringExpression") ?? ""))
            warnings.Add(label + ": dynamic format string differs from the kept measure.");
        if ((GetStringProperty(r, "DetailRowsExpression") ?? "") != (GetStringProperty(kept, "DetailRowsExpression") ?? ""))
            warnings.Add(label + ": detail rows expression differs from the kept measure.");
        if (r.KPI != null)
            warnings.Add(label + ": has a KPI, which is deleted with the measure.");
        if (r.Table != kept.Table)
            warnings.Add(label + ": is in table '" + r.Table.Name + "'; visuals will point to table '" + kept.Table.Name + "'.");
        if (kept.IsHidden && !r.IsHidden)
            warnings.Add(label + ": is visible, the kept measure is hidden.");
    }
    return warnings;
};

// Returns "Merge", "Skip", "Back", "Finish" or "Cancel"
Func<DuplicateGroup, int, string> ShowGroupDialog = (grp, index) =>
{
    string choice = "Cancel";
    using (var form = new Form())
    {
        form.Text = "Duplicate group " + (index + 1) + " of " + groups.Count;
        var screen = BigDialogSize();
        form.Size = new Size(Math.Min(screen.Width, 1300), Math.Min(screen.Height, 720));
        form.MinimumSize = new Size(820, 520);
        form.StartPosition = FormStartPosition.CenterScreen;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 6, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var heading = new Label {
            Text = "Group " + (index + 1) + " of " + groups.Count + ": " + grp.Members.Count + " measures with the same DAX" +
                (grp.Decision != null ? "   (earlier choice: " + (grp.Decision == "Merge" ? "merge" : "do nothing") + ")" : ""),
            AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 6)
        };
        var daxBox = new TextBox {
            Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
            Font = new Font("Consolas", 9), Text = (grp.Keep.Expression ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n")
        };
        var hint = new Label {
            Text = "Keep = the measure that stays. Merge + delete = measures that are repointed to it and deleted. Default keep: most references (PBIR + App + Model).",
            AutoSize = true, MaximumSize = new Size(1200, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 4)
        };

        var grid = new DataGridView {
            Dock = DockStyle.Fill, Font = font, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells, BackgroundColor = SystemColors.Window, MultiSelect = false
        };
        grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Keep", SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Merge + delete", SortMode = DataGridViewColumnSortMode.NotSortable });
        foreach (var header in new[] { "Table", "Measure", "Display folder", "Format string", "Hidden", "PBIR refs", "App refs", "Model refs", "Total", "Description" })
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = header, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });

        var keep = grp.Keep;
        var merge = new HashSet<Measure>(grp.Decision == null ? grp.Members.Where(x => x != keep) : grp.Remove);
        foreach (var x in grp.Members)
        {
            int rowIndex = grid.Rows.Add(false, false, x.Table.Name, x.Name, x.DisplayFolder ?? "", x.FormatString ?? "",
                x.IsHidden ? "Yes" : "No", RefCount(pbirRefs, x), RefCount(appRefs, x), ModelRefs(x), TotalRefs(x), Shorten(x.Description, 80));
            grid.Rows[rowIndex].Tag = x;
        }

        var warningLabel = new Label { AutoSize = true, MaximumSize = new Size(1200, 0), Font = font, ForeColor = Color.DarkRed, Margin = new Padding(0, 8, 0, 0) };

        bool suppress = false;
        Action sync = () =>
        {
            suppress = true;
            foreach (DataGridViewRow row in grid.Rows)
            {
                var x = (Measure)row.Tag;
                bool isKeep = x == keep;
                row.Cells[0].Value = isKeep;
                row.Cells[1].Value = !isKeep && merge.Contains(x);
                row.Cells[1].ReadOnly = isKeep;
                row.DefaultCellStyle.BackColor = isKeep ? Color.Honeydew : (merge.Contains(x) ? Color.MistyRose : SystemColors.Window);
            }
            suppress = false;
            var warnings = GroupWarnings(keep, merge.ToList());
            warningLabel.Text = warnings.Count == 0 ? "" : "Check before merging:\n• " + string.Join("\n• ", warnings);
            daxBox.Text = (keep.Expression ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n");
        };
        grid.CurrentCellDirtyStateChanged += (s, e) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        grid.CellValueChanged += (s, e) =>
        {
            if (suppress || e.RowIndex < 0) return;
            var x = (Measure)grid.Rows[e.RowIndex].Tag;
            if (e.ColumnIndex == 0 && (bool)grid.Rows[e.RowIndex].Cells[0].Value && x != keep)
            {
                merge.Add(keep);   // the previous keep is merged by default
                merge.Remove(x);
                keep = x;
            }
            else if (e.ColumnIndex == 1 && x != keep)
            {
                if ((bool)grid.Rows[e.RowIndex].Cells[1].Value) merge.Add(x); else merge.Remove(x);
            }
            grid.BeginInvoke((Action)(() => sync()));
        };

        var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
        Func<string, string, bool, Button> makeButton = (text, result, bold) =>
        {
            var button = new Button {
                Text = text, AutoSize = true, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0),
                Font = bold ? new Font("Segoe UI", 9, FontStyle.Bold) : font
            };
            button.Click += (s, e) =>
            {
                if (result == "Merge" && merge.Count == 0)
                {
                    MessageBox.Show("Tick at least one measure under Merge + delete, or choose Do nothing.", "Nothing to merge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                choice = result;
                form.Close();
            };
            buttons.Controls.Add(button);
            return button;
        };
        var mergeButton = makeButton("Merge this group", "Merge", true);
        makeButton("Do nothing for this group", "Skip", false);
        makeButton("< Back", "Back", false).Enabled = index > 0;
        makeButton("Finish (skip the remaining groups)", "Finish", false);
        makeButton("Cancel all", "Cancel", false);

        layout.Controls.Add(heading, 0, 0);
        layout.Controls.Add(daxBox, 0, 1);
        layout.Controls.Add(hint, 0, 2);
        layout.Controls.Add(grid, 0, 3);
        layout.Controls.Add(warningLabel, 0, 4);
        layout.Controls.Add(buttons, 0, 5);
        form.Controls.Add(layout);
        form.Shown += (s, e) => { sync(); mergeButton.Focus(); };
        form.ShowDialog();

        if (choice == "Merge")
        {
            grp.Keep = keep;
            grp.Remove = grp.Members.Where(x => x != keep && merge.Contains(x)).ToList();
        }
    }
    return choice;
};

int groupIndex = 0;
while (groupIndex < groups.Count)
{
    var choice = ShowGroupDialog(groups[groupIndex], groupIndex);
    if (choice == "Cancel") { Info("Cancelled - no changes made."); return; }
    if (choice == "Finish") break;
    if (choice == "Back") { groupIndex = Math.Max(0, groupIndex - 1); continue; }
    groups[groupIndex].Decision = choice;
    groupIndex++;
}

var mergedGroups = groups.Where(g => g.Decision == "Merge" && g.Remove.Count > 0).ToList();
if (mergedGroups.Count == 0) { Info("No groups were chosen for merging. No changes made."); return; }

// removed measure -> kept measure
var mergeMap = new Dictionary<Measure, Measure>();
foreach (var grp in mergedGroups)
    foreach (var r in grp.Remove) mergeMap[r] = grp.Keep;
Func<Measure, Measure> ReplaceByMap = msr =>
{
    Measure kept;
    return mergeMap.TryGetValue(msr, out kept) ? kept : null;
};

// ---------------------------------------------------------------------------
// Plan: model DAX that references removed measures
// ---------------------------------------------------------------------------
string[] DaxProperties = { "Expression", "FormatStringExpression", "DetailRowsExpression", "FilterExpression", "TargetExpression", "StatusExpression", "TrendExpression" };

Func<object, string> DescribeObject = o =>
{
    var msr = o as Measure;
    if (msr != null) return "Measure '" + msr.Table.Name + "'[" + msr.Name + "]";
    var cc = o as CalculatedColumn;
    if (cc != null) return "Calculated column '" + cc.Table.Name + "'[" + cc.Name + "]";
    var tbl = o as Table;
    if (tbl != null) return "Calculated table / field parameter '" + tbl.Name + "'";
    var tp = o as TablePermission;
    if (tp != null) return "RLS filter of role '" + tp.Role.Name + "' on '" + tp.Table.Name + "'";
    var ci = o as CalculationItem;
    if (ci != null) return "Calculation item [" + ci.Name + "]";
    var named = o as ITabularNamedObject;
    return o.GetType().Name + (named != null ? " " + named.Name : "");
};

var modelEdits = new List<ModelEdit>();
var dependents = mergeMap.Keys.SelectMany(x => GetReferencedBy(x)).Distinct()
    .Where(o => !(o is Measure && mergeMap.ContainsKey((Measure)o)))
    .ToList();
foreach (var o in dependents)
    foreach (var propertyName in DaxProperties)
    {
        var prop = o.GetType().GetProperty(propertyName);
        if (prop == null || prop.PropertyType != typeof(string) || !prop.CanWrite) continue;
        var before = prop.GetValue(o, null) as string;
        if (string.IsNullOrEmpty(before)) continue;
        var after = RewriteDax(before, ReplaceByMap);
        if (after != before)
            modelEdits.Add(new ModelEdit { Target = o, Property = propertyName, Label = DescribeObject(o) + " (" + propertyName + ")", Before = before, After = after });
    }

// ---------------------------------------------------------------------------
// Plan: PBIR (mutated in memory, written only after confirmation)
// ---------------------------------------------------------------------------
// "Table.Measure" (table names may contain dots) -> the removed measure, or null
Func<string, Measure> RemovedMeasureOfRef = value =>
{
    for (int k = value.IndexOf('.'); k > 0; k = value.IndexOf('.', k + 1))
    {
        var found = FindMeasure(value.Substring(0, k), value.Substring(k + 1));
        if (found != null && mergeMap.ContainsKey(found)) return found;
    }
    return null;
};

// Alias for "table" in a filter's From array; added when the table is not there yet
Func<JArray, string, Dictionary<string, string>, string> AliasFor = (fromArray, table, aliases) =>
{
    foreach (var from in fromArray.OfType<JObject>())
        if (SameName((string)from["Entity"], table) && (string)from["Name"] != null) return (string)from["Name"];
    var used = new HashSet<string>(fromArray.OfType<JObject>().Select(x => (string)x["Name"]).Where(x => x != null), StringComparer.OrdinalIgnoreCase);
    var stem = new string(table.Where(char.IsLetter).Take(1).ToArray()).ToLowerInvariant();
    if (stem == "") stem = "t";
    var alias = stem;
    for (int n = 1; used.Contains(alias); n++) alias = stem + n;
    fromArray.Add(new JObject(new JProperty("Name", alias), new JProperty("Entity", table), new JProperty("Type", 0)));
    aliases[alias] = table;
    return alias;
};

Func<JToken, string, Dictionary<string, string>, JArray, bool, int> mutate = null;
mutate = (token, parentKey, aliasMap, nearestFrom, isExtensions) =>
{
    int changes = 0;
    var obj = token as JObject;
    if (obj != null)
    {
        var aliases = aliasMap;
        var fromArray = obj["From"] as JArray;
        if (fromArray != null)
        {
            aliases = new Dictionary<string, string>(aliasMap, StringComparer.OrdinalIgnoreCase);
            foreach (var from in fromArray.OfType<JObject>())
                if ((string)from["Name"] != null && (string)from["Entity"] != null) aliases[(string)from["Name"]] = (string)from["Entity"];
            nearestFrom = fromArray;
        }

        if (parentKey == "Measure")
        {
            var sourceRef = obj["Expression"]?["SourceRef"] as JObject;
            var table = (string)sourceRef?["Entity"];
            var alias = (string)sourceRef?["Source"];
            if (table == null && alias != null) aliases.TryGetValue(alias, out table);
            var removed = FindMeasure(table, (string)obj["Property"]);
            Measure kept;
            if (removed != null && mergeMap.TryGetValue(removed, out kept))
            {
                if (!SameName(kept.Table.Name, removed.Table.Name))
                {
                    if (sourceRef["Entity"] != null) sourceRef["Entity"] = kept.Table.Name;
                    else if (nearestFrom != null) sourceRef["Source"] = AliasFor(nearestFrom, kept.Table.Name, aliases);
                    else { sourceRef.Remove("Source"); sourceRef["Entity"] = kept.Table.Name; }
                }
                obj["Property"] = kept.Name;
                changes++;
            }
        }

        foreach (var refProperty in new[] { "queryRef", "metadata" })
        {
            var value = obj[refProperty];
            if (value == null || value.Type != JTokenType.String) continue;
            var removed = RemovedMeasureOfRef((string)value);
            if (removed == null) continue;
            var kept = mergeMap[removed];
            obj[refProperty] = kept.Table.Name + "." + kept.Name;
            changes++;
            var native = obj["nativeQueryRef"];
            if (refProperty == "queryRef" && native != null && native.Type == JTokenType.String && SameName((string)native, removed.Name))
            {
                obj["nativeQueryRef"] = kept.Name;
                changes++;
            }
        }

        foreach (var prop in obj.Properties().ToList())
        {
            var current = prop;
            // Property names that are "Table.Measure" keys
            var keyRemoved = prop.Name.Contains(".") ? RemovedMeasureOfRef(prop.Name) : null;
            if (keyRemoved != null)
            {
                var newKey = mergeMap[keyRemoved].Table.Name + "." + mergeMap[keyRemoved].Name;
                if (obj.Property(newKey) == null)
                {
                    current = new JProperty(newKey, prop.Value);
                    prop.Replace(current);
                    changes++;
                }
            }
            // DAX text: report measures (reportExtensions.json) and visual calculations
            if (current.Value.Type == JTokenType.String &&
                ((isExtensions && current.Name == "expression") || (parentKey == "NativeVisualCalculation" && current.Name == "Expression")))
            {
                var before = (string)current.Value;
                var after = RewriteDax(before, ReplaceByMap);
                if (after != before) { current.Value = after; changes++; }
                continue;
            }
            changes += mutate(current.Value, current.Name, aliases, nearestFrom, isExtensions);
        }
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array.ToList()) changes += mutate(item, parentKey, aliasMap, nearestFrom, isExtensions);
    }
    return changes;
};

// Same field twice in one visual role (the visual used both the kept and a removed measure)
Func<JObject, int> DuplicateProjections = docRoot =>
{
    var queryState = docRoot["visual"]?["query"]?["queryState"] as JObject;
    if (queryState == null) return 0;
    int duplicates = 0;
    foreach (var role in queryState.Properties())
    {
        var refs = (role.Value["projections"] as JArray)?.Select(p => (string)p["queryRef"]).Where(x => x != null).ToList();
        if (refs != null) duplicates += refs.Count - refs.Distinct(StringComparer.OrdinalIgnoreCase).Count();
    }
    return duplicates;
};

foreach (var doc in documents)
{
    int duplicatesBefore = DuplicateProjections(doc.Root);
    doc.Changes = mutate(doc.Root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), null, doc.IsExtensions);
    if (doc.Changes > 0 && DuplicateProjections(doc.Root) > duplicatesBefore)
        doc.Warning = "The visual now has the same measure twice in one field well - remove one in Power BI Desktop.";
}
var changedDocuments = documents.Where(d => d.Changes > 0).ToList();

foreach (var doc in appDocuments)
    doc.NewText = RewriteApp(doc, ReplaceByMap);
var changedApps = appDocuments.Where(d => d.NewText != d.Text).ToList();

// ---------------------------------------------------------------------------
// Preview and confirm
// ---------------------------------------------------------------------------
var perspectiveNames = Model.Perspectives.Select(p => p.Name).ToList();
var preview = new System.Data.DataTable();
preview.Columns.Add("Area");
preview.Columns.Add("Object / file");
preview.Columns.Add("Change");
preview.Columns.Add("Before");
preview.Columns.Add("After");
preview.Columns.Add("Warning");

foreach (var grp in mergedGroups)
{
    var keptLabel = "'" + grp.Keep.Table.Name + "'[" + grp.Keep.Name + "]";
    foreach (var r in grp.Remove)
        preview.Rows.Add("Model", "'" + r.Table.Name + "'[" + r.Name + "]", "Delete - merged into " + keptLabel,
            "PBIR " + RefCount(pbirRefs, r) + " / App " + RefCount(appRefs, r) + " / Model " + ModelRefs(r) + " refs", "", "");
    var addPerspectives = perspectiveNames.Where(p => !grp.Keep.InPerspective[p] && grp.Remove.Any(r => r.InPerspective[p])).ToList();
    if (addPerspectives.Count > 0)
        preview.Rows.Add("Model", keptLabel, "Add to perspective(s)", "", string.Join(", ", addPerspectives), "");
    if (string.IsNullOrWhiteSpace(grp.Keep.Description) && grp.Remove.Any(r => !string.IsNullOrWhiteSpace(r.Description)))
        preview.Rows.Add("Model", keptLabel, "Copy description", "", Shorten(grp.Remove.First(r => !string.IsNullOrWhiteSpace(r.Description)).Description, 120), "");
    var warnings = GroupWarnings(grp.Keep, grp.Remove);
    if (warnings.Count > 0)
        preview.Rows.Add("Model", keptLabel, "Differences accepted", "", "", string.Join(" | ", warnings));
}
foreach (var edit in modelEdits)
    preview.Rows.Add("Model DAX", edit.Label, "Rewrite references", Shorten(edit.Before, 300), Shorten(edit.After, 300), "");
foreach (var doc in changedDocuments)
    preview.Rows.Add("PBIR", doc.ReportName + " | " + doc.File.Substring(doc.DefinitionFolder.Length).TrimStart('\\', '/'),
        doc.Changes + " reference value(s)", "", "", doc.Warning ?? "");
foreach (var doc in changedApps)
{
    var rel = doc.AppName + " | " + doc.File.Substring(doc.AppFolder.Length).TrimStart('\\', '/');
    var oldLines = doc.Text.Split('\n');
    var newLines = doc.NewText.Split('\n');
    for (int k = 0; k < Math.Min(oldLines.Length, newLines.Length); k++)
        if (oldLines[k] != newLines[k])
            preview.Rows.Add("App (experimental)", rel, "Line " + (k + 1), oldLines[k].Trim(), newLines[k].Trim(), "Not thoroughly tested - review");
}

int removeCount = mergeMap.Count;
if (!ShowGrid("Preview merge of duplicate measures",
        "Merge " + mergedGroups.Count + " group(s): delete " + removeCount + " measure(s), rewrite " + modelEdits.Count + " model DAX expression(s), update " +
        changedDocuments.Count + " PBIR file(s)" + (selectedApps.Count > 0 ? " and " + changedApps.Count + " app file(s) (experimental)" : "") + ". Continue to confirm.",
        preview, "Continue"))
{ Info("Cancelled - no changes made."); return; }

var notUpdatedWarning = (scanPbir || scanApps) ? "" :
    "\n\nNo reports or apps are updated - visuals that use the removed measures will break.";
if (MessageBox.Show(
        "Apply the merge?\n\n" +
        "• Model: delete " + removeCount + " measure(s), rewrite " + modelEdits.Count + " DAX expression(s) (undo with Ctrl+Z).\n" +
        "• PBIR: write " + changedDocuments.Count + " file(s).\n" +
        (selectedApps.Count > 0 ? "• Fabric Apps (experimental, not thoroughly tested): write " + changedApps.Count + " file(s).\n" : "") +
        "\nFile changes can NOT be undone in Tabular Editor - make sure they are committed to git and Power BI Desktop is closed." + notUpdatedWarning,
        "Confirm merge", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
{ Info("Cancelled - no changes made."); return; }

// ---------------------------------------------------------------------------
// Apply: model first (undoable), then files
// ---------------------------------------------------------------------------
foreach (var edit in modelEdits)
    edit.Target.GetType().GetProperty(edit.Property).SetValue(edit.Target, edit.After, null);

foreach (var grp in mergedGroups)
{
    foreach (var p in perspectiveNames)
        if (!grp.Keep.InPerspective[p] && grp.Remove.Any(r => r.InPerspective[p])) grp.Keep.InPerspective[p] = true;
    if (string.IsNullOrWhiteSpace(grp.Keep.Description))
    {
        var withDescription = grp.Remove.FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.Description));
        if (withDescription != null) grp.Keep.Description = withDescription.Description;
    }
}
foreach (var r in mergeMap.Keys.ToList()) r.Delete();

int filesWritten = 0;
foreach (var doc in changedDocuments)
{
    try
    {
        File.WriteAllText(doc.File, doc.Root.ToString(Newtonsoft.Json.Formatting.Indented));
        filesWritten++;
    }
    catch (Exception ex)
    {
        Error("Stopped while writing:\n" + doc.File + "\n\nAlready done: model merge (undo with Ctrl+Z) and " + filesWritten + " file(s) written.\n\nError: " + ex.Message);
    }
}
int appFilesWritten = 0;
foreach (var doc in changedApps)
{
    try
    {
        File.WriteAllText(doc.File, doc.NewText);
        appFilesWritten++;
    }
    catch (Exception ex)
    {
        Error("Stopped while writing:\n" + doc.File + "\n\nAlready done: model merge (undo with Ctrl+Z), " + filesWritten + " PBIR file(s) and " + appFilesWritten + " app file(s) written.\n\nError: " + ex.Message);
    }
}

int duplicateWarnings = changedDocuments.Count(d => d.Warning != null);
Info("Done. Merged " + mergedGroups.Count + " group(s): deleted " + removeCount + " measure(s), rewrote " + modelEdits.Count + " model DAX expression(s), wrote " +
    filesWritten + " PBIR file(s)" + (selectedApps.Count > 0 ? " and " + appFilesWritten + " app file(s) (experimental - test the apps)" : "") + "." +
    (duplicateWarnings > 0 ? "\n\n" + duplicateWarnings + " visual(s) now show the same measure twice - see the preview warnings." : "") +
    "\n\nNext:\n1. Save the model in Tabular Editor (and deploy it if reports are live-connected).\n" +
    "2. Reopen the reports in Power BI Desktop (or use Desktop Bridge) and check the changed visuals.");
