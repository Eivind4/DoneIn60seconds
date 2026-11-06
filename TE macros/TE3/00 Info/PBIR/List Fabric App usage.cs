// ============================================================================
// 21 - List Fabric App usage.cs
//
// Read-only OVERVIEW of what Fabric Apps (Rayfin projects) query from a
// semantic model: one row per model measure / column, query measure, query
// alias or reference that is not in the model, with the apps and files that
// use it. Nothing is changed.
//
// Resolve references against (first that applies):
//   - The model open in Tabular Editor (default when one with tables is open).
//   - A model definition on disk (TMDL folder or model.bim).
//   - The app's own model read online (fabric.yaml connection).
//   - Nothing: references are listed unverified.
// So it also works when the model cannot be opened in Tabular Editor (for
// example an Enterprise-tier model with a Business Edition licence). If Tabular
// Editor refuses to run a script with no model loaded, open any small local
// model first and untick "Resolve against the model open in Tabular Editor".
//
// (Replaces the former 21 that needed the open model and 21b, the offline
// test version. For usage combined with PBIR reports, see macro 22.)
//
// Inputs (one dialog):
//   1. Apps folder        - every rayfin/rayfin.yml below it is a Fabric App
//                           (node_modules, dist and .agents are skipped).
//   2. Model filter       - optional. Semantic model item id, workspace id or
//                           connector alias from rayfin.yml. Only apps with a
//                           matching fabric-semanticmodel connector are scanned.
//                           Leave empty to scan every app.
//   3. Model definition   - optional. A .SemanticModel folder, its definition
//                           folder (TMDL) or a model.bim file. When given,
//                           references are resolved against it.
//   4. Verify online      - optional, used when no model definition is given.
//                           Reads tables, columns and measures (INFO.VIEW.*) of
//                           the model each app connects to, through the app's
//                           own fabric.yaml connection: "npx fabric-app-data
//                           query" run in the app folder. Needs Node.js, the
//                           app's node_modules (npm install) and a Fabric
//                           sign-in that can read the model. Apps whose model is
//                           only declared in rayfin.yml are listed unverified.
//                           When nothing verifies, references are listed
//                           unverified:
//                             Table[Name]  -> "Column (unverified)"
//                             [Name]       -> "Measure (unverified)", or "Query alias"
//                                             when the query names a result column
//                                             "Name" itself (SELECTCOLUMNS, ROW, ...)
//
// DAX parsing: .dax files are fully parsed,
// .ts/.tsx files only contribute names that exist in the model definition
// (so without a definition, .ts files are skipped).
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

class ConnectorInfo
{
    public string Name;
    public string WorkspaceId;
    public string ItemId;
    public string Source;    // rayfin.yml or fabric.yaml
}

class AppInfo
{
    public string AppRoot;
    public string AppName;
    public List<ConnectorInfo> Connectors = new List<ConnectorInfo>();
    public string ConnectorSummary
    {
        get
        {
            return Connectors.Count == 0
                ? "no semantic-model connector"
                : string.Join(", ", Connectors.Select(c => c.Name + " (model " + (c.ItemId ?? "?") + ")"));
        }
    }
}

class ModelObject
{
    public string Kind;      // Measure or Column
    public string Table;
    public string Name;
    public string DisplayFolder = "";
    public bool Hidden;
}

class ModelCatalog
{
    public string Name;
    public HashSet<string> Tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, ModelObject> Columns = new Dictionary<string, ModelObject>(StringComparer.OrdinalIgnoreCase); // "table|column"
    public Dictionary<string, ModelObject> Measures = new Dictionary<string, ModelObject>(StringComparer.OrdinalIgnoreCase);
}

class AppReference
{
    public string App;
    public string Kind;
    public string Table;
    public string Name;
    public string File;
    public ModelObject Object;
    public string Key { get { return Kind + "|" + Table + "|" + Name; } }
}

string[] SkippedFolders = { "node_modules", "dist", ".agents", ".git", ".temp", ".vite", "coverage" };

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

// Top-level "connectors:" block of rayfin.yml (services.connectors is only the on/off switch).
Func<string, List<ConnectorInfo>> ReadConnectors = rayfinYml =>
{
    var list = new List<ConnectorInfo>();
    bool inBlock = false;
    ConnectorInfo current = null;
    string type = null;
    Action flush = () =>
    {
        if (current != null && string.Equals(type, "fabric-semanticmodel", StringComparison.OrdinalIgnoreCase))
            list.Add(current);
        current = null;
        type = null;
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
        Func<string, string> value = prefix => line.Substring(prefix.Length).Trim().Trim('"', '\'');
        if (line.StartsWith("- name:")) { flush(); current = new ConnectorInfo { Name = value("- name:"), Source = "rayfin.yml" }; }
        else if (current == null) continue;
        else if (line.StartsWith("type:")) type = value("type:");
        else if (line.StartsWith("workspaceId:")) current.WorkspaceId = value("workspaceId:");
        else if (line.StartsWith("itemId:")) current.ItemId = value("itemId:");
    }
    if (inBlock) flush();
    return list;
};

// Semantic models declared in the app's fabric.yaml (fabric-app-data, used by the
// Data App template) for the active profile: profiles.<p>.semanticModels.<alias>.
Func<string, List<ConnectorInfo>> ReadFabricYaml = appRoot =>
{
    var list = new List<ConnectorInfo>();
    var file = new[] { "fabric.yaml", "fabric.yml" }.Select(n => Path.Combine(appRoot, n)).FirstOrDefault(File.Exists);
    if (file == null) return list;
    var lines = File.ReadAllLines(file);
    string active = lines.Where(l => l.StartsWith("activeProfile:"))
        .Select(l => l.Substring(14).Trim().Trim('"', '\'')).FirstOrDefault();
    bool inProfiles = false, inModels = false;
    int profileIndent = -1, modelsIndent = -1, aliasIndent = -1;
    string profile = null;
    ConnectorInfo current = null;
    Action flush = () =>
    {
        if (current != null && (string.IsNullOrEmpty(active) || profile == active)) list.Add(current);
        current = null;
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
        if (aliasIndent < 0 || indent <= aliasIndent) { flush(); aliasIndent = indent; current = new ConnectorInfo { Name = key, Source = "fabric.yaml" }; continue; }
        if (key == "workspaceId") current.WorkspaceId = val;
        else if (key == "itemId") current.ItemId = val;
    }
    flush();
    return list;
};

// Runs one DAX query through the app's own fabric-app-data CLI (same connection and
// sign-in as "npm run dev") and returns the parsed JSON result.
Func<string, string, string, JObject> RunAppQuery = (appRoot, alias, dax) =>
{
    var daxFile = Path.Combine(Path.GetTempPath(), "te-app-catalog-" + Guid.NewGuid().ToString("N") + ".dax");
    File.WriteAllText(daxFile, dax);
    try
    {
        var psi = new System.Diagnostics.ProcessStartInfo("cmd.exe",
            "/d /c npx --no-install fabric-app-data query \"" + alias + "\" --limit 1000 --file \"" + daxFile + "\"")
        {
            WorkingDirectory = appRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        using (var p = System.Diagnostics.Process.Start(psi))
        {
            var stderr = p.StandardError.ReadToEndAsync();
            var stdout = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(180000)) { try { p.Kill(); } catch { } throw new Exception("fabric-app-data query timed out after 3 minutes."); }
            var start = stdout.IndexOf('{');
            if (start < 0)
            {
                var err = (stderr.Result ?? "").Trim();
                throw new Exception(err.Length > 0 ? err.Substring(0, Math.Min(err.Length, 600)) : "No output from fabric-app-data (is node_modules installed? run npm install in the app).");
            }
            var result = JObject.Parse(stdout.Substring(start));
            if ((string)result["status"] != "success")
                throw new Exception((string)result["error"]?["message"] ?? "Query failed.");
            return result;
        }
    }
    finally { try { File.Delete(daxFile); } catch { } }
};

// Builds a catalog from INFO.VIEW.TABLES/COLUMNS/MEASURES of the app's model. Pages by
// [ID] because the CLI returns at most 1000 rows per query.
Func<string, string, ModelCatalog> LoadOnlineCatalog = (appRoot, alias) =>
{
    var catalog = new ModelCatalog { Name = alias + " (online)" };
    foreach (var kind in new[] { "Table", "Column", "Measure" })
    {
        var view = kind == "Table" ? "TABLES" : kind == "Column" ? "COLUMNS" : "MEASURES";
        var tableExpr = kind == "Table" ? "[Name]" : "[Table]";
        var folderExpr = kind == "Table" ? "\"\"" : "[DisplayFolder]";
        long lastId = -1;
        while (true)
        {
            var dax = "EVALUATE TOPN ( 1000, FILTER ( SELECTCOLUMNS ( INFO.VIEW." + view + " (), \"ID\", [ID], \"Table\", " + tableExpr +
                ", \"Name\", [Name], \"Folder\", " + folderExpr + ", \"Hidden\", [IsHidden] ), [ID] > " + lastId + " ), [ID], ASC ) ORDER BY [ID]";
            var rows = RunAppQuery(appRoot, alias, dax)["table"]?["rows"] as JArray ?? new JArray();
            foreach (var row in rows)
            {
                lastId = Math.Max(lastId, (long)row[0]);
                var table = (string)row[1];
                var name = (string)row[2];
                if (kind == "Table") { catalog.Tables.Add(table); continue; }
                if (kind == "Column" && name.StartsWith("RowNumber-")) continue;
                var o = new ModelObject { Kind = kind, Table = table, Name = name, DisplayFolder = (string)row[3] ?? "", Hidden = (bool?)row[4] ?? false };
                if (kind == "Column") catalog.Columns[table + "|" + name] = o;
                else if (!catalog.Measures.ContainsKey(name)) catalog.Measures.Add(name, o);
            }
            if (rows.Count < 1000) break;
        }
    }
    return catalog;
};

// Reads a TMDL name after a keyword: 'Quoted name' or bare-name (up to whitespace or '=').
Func<string, string> ReadTmdlName = rest =>
{
    rest = rest.TrimStart();
    if (rest.StartsWith("'"))
    {
        var sb = new StringBuilder();
        for (int i = 1; i < rest.Length; i++)
        {
            if (rest[i] == '\'' && i + 1 < rest.Length && rest[i + 1] == '\'') { sb.Append('\''); i++; continue; }
            if (rest[i] == '\'') break;
            sb.Append(rest[i]);
        }
        return sb.ToString();
    }
    var m = Regex.Match(rest, @"^[^\s=]+");
    return m.Success ? m.Value : rest.Trim();
};

// Builds a catalog of tables, columns and measures from TMDL or model.bim on disk.
Func<string, ModelCatalog> LoadModelDefinition = path =>
{
    var catalog = new ModelCatalog();

    if (File.Exists(path) && path.EndsWith(".bim", StringComparison.OrdinalIgnoreCase))
    {
        var root = JObject.Parse(File.ReadAllText(path));
        var model = root["model"] ?? root;
        catalog.Name = (string)root["name"] ?? Path.GetFileNameWithoutExtension(path);
        foreach (var t in model["tables"] ?? new JArray())
        {
            var table = (string)t["name"];
            catalog.Tables.Add(table);
            foreach (var c in t["columns"] ?? new JArray())
            {
                if ((string)c["type"] == "rowNumber") continue;
                var o = new ModelObject { Kind = "Column", Table = table, Name = (string)c["name"], DisplayFolder = (string)c["displayFolder"] ?? "", Hidden = (bool?)c["isHidden"] ?? false };
                catalog.Columns[table + "|" + o.Name] = o;
            }
            foreach (var ms in t["measures"] ?? new JArray())
            {
                var o = new ModelObject { Kind = "Measure", Table = table, Name = (string)ms["name"], DisplayFolder = (string)ms["displayFolder"] ?? "", Hidden = (bool?)ms["isHidden"] ?? false };
                if (!catalog.Measures.ContainsKey(o.Name)) catalog.Measures.Add(o.Name, o);
            }
        }
        return catalog;
    }

    // TMDL: accept the .SemanticModel folder, its definition folder or the tables folder.
    var dir = path;
    if (Directory.Exists(Path.Combine(dir, "definition"))) dir = Path.Combine(dir, "definition");
    var tablesDir = Directory.Exists(Path.Combine(dir, "tables")) ? Path.Combine(dir, "tables") : dir;

    var platform = new[] { Path.Combine(path, ".platform"), Path.Combine(Path.GetDirectoryName(dir) ?? dir, ".platform") }.FirstOrDefault(File.Exists);
    if (platform != null)
    {
        try { catalog.Name = (string)JObject.Parse(File.ReadAllText(platform))["metadata"]?["displayName"]; } catch { }
    }
    if (string.IsNullOrEmpty(catalog.Name))
    {
        var folderName = Path.GetFileName(path.TrimEnd('\\', '/'));
        catalog.Name = folderName.EndsWith(".SemanticModel") ? folderName.Substring(0, folderName.Length - ".SemanticModel".Length) : folderName;
    }

    foreach (var file in Directory.GetFiles(tablesDir, "*.tmdl"))
    {
        string table = null;
        ModelObject current = null;
        foreach (var raw in File.ReadAllLines(file))
        {
            if (raw.StartsWith("table "))
            {
                table = ReadTmdlName(raw.Substring(6));
                catalog.Tables.Add(table);
                current = null;
                continue;
            }
            if (table == null) continue;

            // Object declarations are indented one tab; their properties two tabs.
            var decl = Regex.Match(raw, @"^\t(column|measure)\s+(.*)$");
            if (decl.Success)
            {
                var kind = decl.Groups[1].Value == "column" ? "Column" : "Measure";
                current = new ModelObject { Kind = kind, Table = table, Name = ReadTmdlName(decl.Groups[2].Value) };
                if (kind == "Column") catalog.Columns[table + "|" + current.Name] = current;
                else if (!catalog.Measures.ContainsKey(current.Name)) catalog.Measures.Add(current.Name, current);
                continue;
            }
            if (Regex.IsMatch(raw, @"^\t\S")) { current = null; continue; } // another object type (hierarchy, partition, ...)

            if (current != null && raw.StartsWith("\t\t"))
            {
                var prop = raw.Trim();
                if (prop.StartsWith("displayFolder:")) current.DisplayFolder = prop.Substring(14).Trim().Trim('\'', '"');
                else if (prop == "isHidden" || prop == "isHidden: true") current.Hidden = true;
            }
        }
    }
    return catalog;
};

// ---- DAX parsing (same rules as macros 19, 22 and 23) ----------------------------------
Func<string, string> StripDaxCommentsAndStrings = dax =>
{
    var sb = new StringBuilder(dax.Length);
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

var QualifiedRef = new Regex(@"(?<![\w\].])('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]", RegexOptions.Compiled);
var StringLiteral = new Regex("\"((?:[^\"]|\"\")*)\"", RegexOptions.Compiled);
var BareRef = new Regex(@"\[((?:[^\]]|\]\])+)\]", RegexOptions.Compiled);
var DefineMeasure = new Regex(@"\bMEASURE\s+('(?:[^']|'')+'|[A-Za-z_][\w]*)\s*\[((?:[^\]]|\]\])+)\]\s*=", RegexOptions.Compiled | RegexOptions.IgnoreCase);
var DaxKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
    "BY", "RETURN", "IN", "VAR", "EVALUATE", "DEFINE", "ORDER", "ASC", "DESC", "AND", "OR", "NOT", "START", "AT", "MEASURE", "COLUMN", "TABLE"
};

Func<string, string> UnquoteTable = t => t.StartsWith("'") ? t.Substring(1, t.Length - 2).Replace("''", "'") : t;
Func<string, string> UnescapeName = n => n.Replace("]]", "]");

// The model open in Tabular Editor, when there is one with tables. Read in a try:
// this macro must also run when the real model cannot be opened (see header).
ModelCatalog openModelCatalog = null;
try
{
    if (Model != null && Model.Tables.Any())
    {
        string dbName = null;
        try { dbName = Model.Database?.Name; } catch { }
        var c = new ModelCatalog { Name = (string.IsNullOrEmpty(dbName) ? "model" : dbName) + " (open in Tabular Editor)" };
        foreach (var t in Model.Tables)
        {
            c.Tables.Add(t.Name);
            foreach (var col in t.Columns)
            {
                if (col.Type.ToString() == "RowNumber") continue;
                c.Columns[t.Name + "|" + col.Name] = new ModelObject { Kind = "Column", Table = t.Name, Name = col.Name, DisplayFolder = col.DisplayFolder ?? "", Hidden = col.IsHidden };
            }
            foreach (var ms in t.Measures)
                if (!c.Measures.ContainsKey(ms.Name))
                    c.Measures.Add(ms.Name, new ModelObject { Kind = "Measure", Table = t.Name, Name = ms.Name, DisplayFolder = ms.DisplayFolder ?? "", Hidden = ms.IsHidden });
        }
        openModelCatalog = c;
    }
}
catch { openModelCatalog = null; }

// ---------------------------------------------------------------------------
// Dialog: inputs
// ---------------------------------------------------------------------------
string appsFolder = null, modelFilter = null, modelPath = null;
bool verifyOnline = false, useOpenModel = false;
using (var form = new Form())
{
    form.Text = "Fabric App usage";
    form.Width = 760;
    form.AutoSize = true;
    form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.MaximizeBox = form.MinimizeBox = false;
    form.StartPosition = FormStartPosition.CenterScreen;
    form.Font = new Font("Segoe UI", 9);

    var grid = new TableLayoutPanel { AutoSize = true, ColumnCount = 3, Padding = new Padding(12), Dock = DockStyle.Fill };
    grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 460));
    grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

    Func<string, string, TextBox> addRow = (label, hint) =>
    {
        grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 8, 8, 0) });
        var box = new TextBox { Width = 450, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        grid.Controls.Add(box);
        return box;
    };

    var appsBox = addRow("Apps folder", "");
    var appsBrowse = new Button { Text = "Browse…", AutoSize = true };
    appsBrowse.Click += (s, e) =>
    {
        using (var fbd = new FolderBrowserDialog { Description = "Folder containing one or more Fabric Apps (rayfin/rayfin.yml)" })
            if (fbd.ShowDialog() == DialogResult.OK) appsBox.Text = fbd.SelectedPath;
    };
    grid.Controls.Add(appsBrowse);

    var filterBox = addRow("Model filter (optional)", "");
    grid.Controls.Add(new Label { Text = "item id, workspace id or connector alias", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 0) });

    var modelBox = addRow("Model definition (optional)", "");
    var modelButtons = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0) };
    var tmdlBrowse = new Button { Text = "TMDL folder…", AutoSize = true };
    tmdlBrowse.Click += (s, e) =>
    {
        using (var fbd = new FolderBrowserDialog { Description = "A .SemanticModel folder or its definition folder (TMDL)" })
            if (fbd.ShowDialog() == DialogResult.OK) modelBox.Text = fbd.SelectedPath;
    };
    var bimBrowse = new Button { Text = "model.bim…", AutoSize = true };
    bimBrowse.Click += (s, e) =>
    {
        using (var ofd = new OpenFileDialog { Filter = "Tabular model (*.bim)|*.bim|All files (*.*)|*.*" })
            if (ofd.ShowDialog() == DialogResult.OK) modelBox.Text = ofd.FileName;
    };
    modelButtons.Controls.Add(tmdlBrowse);
    modelButtons.Controls.Add(bimBrowse);
    grid.Controls.Add(modelButtons);

    var chkOnline = new CheckBox {
        Text = "Verify online when no model definition is given: read each app's model through its fabric.yaml connection (npx fabric-app-data, needs npm install and Fabric sign-in)",
        Checked = true, AutoSize = true, MaximumSize = new Size(700, 0), Margin = new Padding(0, 10, 0, 4)
    };
    grid.Controls.Add(chkOnline);
    grid.SetColumnSpan(chkOnline, 3);

    var chkOpenModel = new CheckBox {
        Text = openModelCatalog != null
            ? "Resolve against the model open in Tabular Editor: " + openModelCatalog.Name + " (model definition and online verification are then not used)"
            : "Resolve against the model open in Tabular Editor (no model with tables is open)",
        Checked = openModelCatalog != null, Enabled = openModelCatalog != null,
        AutoSize = true, MaximumSize = new Size(700, 0), Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 4, 0, 4)
    };
    grid.Controls.Add(chkOpenModel);
    grid.SetColumnSpan(chkOpenModel, 3);
    Action syncModelInputs = () =>
    {
        var other = !chkOpenModel.Checked;
        modelBox.Enabled = tmdlBrowse.Enabled = bimBrowse.Enabled = chkOnline.Enabled = other;
    };
    chkOpenModel.CheckedChanged += (s, e) => syncModelInputs();
    syncModelInputs();

    var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
    var ok = new Button { Text = "Scan", DialogResult = DialogResult.OK, AutoSize = true };
    var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
    buttons.Controls.Add(ok);
    buttons.Controls.Add(cancel);
    grid.Controls.Add(buttons);
    grid.SetColumnSpan(buttons, 3);

    form.Controls.Add(grid);
    form.AcceptButton = ok;
    form.CancelButton = cancel;

    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled."); return; }
    appsFolder = appsBox.Text.Trim().Trim('"');
    modelFilter = filterBox.Text.Trim();
    modelPath = modelBox.Text.Trim().Trim('"');
    verifyOnline = chkOnline.Checked;
    useOpenModel = chkOpenModel.Checked && openModelCatalog != null;
}

if (!Directory.Exists(appsFolder)) { Info("Apps folder not found:\n" + appsFolder); return; }
if (useOpenModel) { modelPath = ""; verifyOnline = false; }
if (modelPath.Length > 0 && !Directory.Exists(modelPath) && !File.Exists(modelPath)) { Info("Model definition not found:\n" + modelPath); return; }

ModelCatalog catalog = useOpenModel ? openModelCatalog : null;
if (modelPath.Length > 0)
{
    catalog = LoadModelDefinition(modelPath);
    if (catalog.Tables.Count == 0) { Info("No tables found in the model definition:\n" + modelPath); return; }
}

// ---------------------------------------------------------------------------
// Find apps and apply the model filter
// ---------------------------------------------------------------------------
var apps = EnumerateSourceFiles(appsFolder)
    .Where(f => string.Equals(Path.GetFileName(f), "rayfin.yml", StringComparison.OrdinalIgnoreCase)
             && string.Equals(Path.GetFileName(Path.GetDirectoryName(f)), "rayfin", StringComparison.OrdinalIgnoreCase))
    .Select(f =>
    {
        var root = Path.GetDirectoryName(Path.GetDirectoryName(f));
        return new AppInfo { AppRoot = root, AppName = Path.GetFileName(root), Connectors = ReadConnectors(f).Concat(ReadFabricYaml(root)).ToList() };
    })
    .OrderBy(a => a.AppName)
    .ToList();

if (apps.Count == 0) { Info("No Fabric Apps (rayfin/rayfin.yml) were found below:\n" + appsFolder); return; }

Func<ConnectorInfo, bool> MatchesFilter = c =>
    modelFilter.Length == 0
    || string.Equals(c.ItemId, modelFilter, StringComparison.OrdinalIgnoreCase)
    || string.Equals(c.WorkspaceId, modelFilter, StringComparison.OrdinalIgnoreCase)
    || string.Equals(c.Name, modelFilter, StringComparison.OrdinalIgnoreCase);

var selectedApps = apps.Where(a => modelFilter.Length == 0 || a.Connectors.Any(MatchesFilter)).ToList();
if (selectedApps.Count == 0)
{
    Info("No app has a semantic-model connector matching \"" + modelFilter + "\".\n\nApps found:\n" +
         string.Join("\n", apps.Select(a => "• " + a.AppName + " → " + a.ConnectorSummary)));
    return;
}

// ---------------------------------------------------------------------------
// Scan
// ---------------------------------------------------------------------------
var references = new List<AppReference>();
int daxFiles = 0, codeFiles = 0;
string unverifiedColumn = "Column (unverified)";
string unverifiedBare = "Measure (unverified)";

// The catalog references are resolved against: the model definition for every app,
// or the app's own model read online, or none (unverified).
ModelCatalog activeCatalog = null;
var verification = new List<string>();
var onlineCatalogs = new Dictionary<string, ModelCatalog>(StringComparer.OrdinalIgnoreCase);

Func<string, string, ModelObject> FindColumn = (table, name) =>
{
    ModelObject o;
    return activeCatalog != null && activeCatalog.Columns.TryGetValue(table + "|" + name, out o) ? o : null;
};
Func<string, ModelObject> FindMeasure = name =>
{
    ModelObject o;
    return activeCatalog != null && activeCatalog.Measures.TryGetValue(name, out o) ? o : null;
};

foreach (var app in selectedApps)
{
    var files = EnumerateSourceFiles(app.AppRoot).ToList();

    activeCatalog = catalog;
    if (catalog != null)
        verification.Add(app.AppName + ": verified against model definition " + catalog.Name);
    else if (verifyOnline)
    {
        var online = app.Connectors.FirstOrDefault(c => c.Source == "fabric.yaml" && MatchesFilter(c));
        if (online == null)
            verification.Add(app.AppName + ": unverified - no fabric.yaml semantic model to read online");
        else
        {
            var cacheKey = app.AppRoot + "|" + online.Name;
            try
            {
                if (!onlineCatalogs.ContainsKey(cacheKey)) onlineCatalogs[cacheKey] = LoadOnlineCatalog(app.AppRoot, online.Name);
                activeCatalog = onlineCatalogs[cacheKey];
                verification.Add(app.AppName + ": verified online via '" + online.Name + "' (model " + (online.ItemId ?? "?") + "), " +
                    activeCatalog.Tables.Count + " tables, " + activeCatalog.Columns.Count + " columns, " + activeCatalog.Measures.Count + " measures");
            }
            catch (Exception ex)
            {
                verification.Add(app.AppName + ": unverified - online read via '" + online.Name + "' failed: " + ex.Message);
            }
        }
    }
    else
        verification.Add(app.AppName + ": unverified - no model definition and online verification off");
    Action<string, string, string, string, ModelObject> add = (kind, table, name, rel, obj) =>
        references.Add(new AppReference { App = app.AppName, Kind = kind, Table = table, Name = name, File = rel, Object = obj });

    foreach (var file in files.Where(f => f.EndsWith(".dax", StringComparison.OrdinalIgnoreCase)))
    {
        daxFiles++;
        var rel = file.Substring(app.AppRoot.Length).TrimStart('\\', '/');
        var rawDax = File.ReadAllText(file);
        var dax = StripDaxCommentsAndStrings(rawDax);

        // Result-column names created in the query: "Week" in SELECTCOLUMNS / ADDCOLUMNS /
        // ROW / GROUPBY / SUMMARIZECOLUMNS. A bare [Week] that matches one is an alias.
        var queryAliases = new HashSet<string>(
            StringLiteral.Matches(rawDax).Cast<Match>().Select(x => x.Groups[1].Value.Replace("\"\"", "\"")),
            StringComparer.OrdinalIgnoreCase);

        var queryMeasures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match m in DefineMeasure.Matches(dax))
        {
            var name = UnescapeName(m.Groups[2].Value);
            queryMeasures.Add(name);
            add("Query measure", UnquoteTable(m.Groups[1].Value), name, rel, null);
        }

        var remaining = new StringBuilder(dax);
        foreach (Match m in QualifiedRef.Matches(dax))
        {
            var table = UnquoteTable(m.Groups[1].Value);
            // A keyword (ORDER BY [Week]) or, with a definition loaded, any non-table:
            // leave the bracket name to the bare pass.
            if (activeCatalog != null ? !activeCatalog.Tables.Contains(table) : DaxKeywords.Contains(table)) continue;
            var name = UnescapeName(m.Groups[2].Value);
            for (int k = m.Index; k < m.Index + m.Length; k++) remaining[k] = ' ';
            if (queryMeasures.Contains(name)) continue; // DEFINE MEASURE header

            if (activeCatalog == null) { add(unverifiedColumn, table, name, rel, null); continue; }
            var col = FindColumn(table, name);
            var meas = col == null ? FindMeasure(name) : null;
            if (col != null) add("Column", col.Table, col.Name, rel, col);
            else if (meas != null) add("Measure", meas.Table, meas.Name, rel, meas);
            else add("Not in model", table, name, rel, null);
        }

        foreach (Match m in BareRef.Matches(remaining.ToString()))
        {
            var name = UnescapeName(m.Groups[1].Value);
            if (queryMeasures.Contains(name)) continue;
            var meas = FindMeasure(name);
            if (meas != null) add("Measure", meas.Table, meas.Name, rel, meas);
            else if (queryAliases.Contains(name)) add("Query alias", "", name, rel, null);
            else if (activeCatalog == null) add(unverifiedBare, "", name, rel, null);
            else add("Not in model", "", name, rel, null);
        }
    }

    // .ts/.tsx only contribute names that exist in the model (definition or online).
    if (activeCatalog == null) continue;
    foreach (var file in files.Where(f =>
        (f.EndsWith(".ts", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".tsx", StringComparison.OrdinalIgnoreCase))
        && !Regex.IsMatch(Path.GetFileName(f), @"\.(spec|test)\.tsx?$", RegexOptions.IgnoreCase)
        && !f.EndsWith(".d.ts", StringComparison.OrdinalIgnoreCase)))
    {
        codeFiles++;
        var rel = file.Substring(app.AppRoot.Length).TrimStart('\\', '/');
        var code = File.ReadAllText(file);
        foreach (Match m in BareRef.Matches(code))
        {
            var meas = FindMeasure(UnescapeName(m.Groups[1].Value));
            if (meas != null) add("Measure", meas.Table, meas.Name, rel, meas);
        }
        foreach (Match m in QualifiedRef.Matches(code))
        {
            var col = FindColumn(UnquoteTable(m.Groups[1].Value), UnescapeName(m.Groups[2].Value));
            if (col != null) add("Column", col.Table, col.Name, rel, col);
        }
    }
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------
var kindOrder = new List<string> { "Measure", "Column", unverifiedBare, unverifiedColumn, "Query measure", "Query alias", "Not in model" };

var dt = new System.Data.DataTable();
dt.Columns.Add("Kind");
dt.Columns.Add("Table");
dt.Columns.Add("Name");
dt.Columns.Add("Display folder");
dt.Columns.Add("Hidden");
dt.Columns.Add("Apps");
dt.Columns.Add("References", typeof(int));
dt.Columns.Add("Files");

foreach (var g in references.GroupBy(r => r.Key, StringComparer.OrdinalIgnoreCase)
    .OrderBy(g => kindOrder.IndexOf(g.First().Kind))
    .ThenBy(g => g.First().Table)
    .ThenBy(g => g.First().Name))
{
    var first = g.First();
    dt.Rows.Add(
        first.Kind,
        first.Table,
        first.Name,
        first.Object != null ? first.Object.DisplayFolder : "",
        first.Object != null ? first.Object.Hidden.ToString() : "",
        string.Join(", ", g.Select(r => r.App).Distinct().OrderBy(a => a)),
        g.Count(),
        string.Join("; ", g.Select(r => r.App + "/" + r.File.Replace('\\', '/')).Distinct().OrderBy(f => f)));
}

var summary = new StringBuilder();
summary.AppendLine("Scanned " + selectedApps.Count + " of " + apps.Count + " app(s): " + daxFiles + " .dax and " + codeFiles + " .ts/.tsx file(s).");
summary.AppendLine("Model filter: " + (modelFilter.Length == 0 ? "(none - all apps)" : modelFilter));
summary.AppendLine();
foreach (var a in selectedApps) summary.AppendLine("• " + a.AppName + " → " + a.ConnectorSummary);
summary.AppendLine();
foreach (var v in verification) summary.AppendLine("• " + v);
summary.AppendLine();
var usedMeasures = references.Where(r => r.Kind == "Measure").Select(r => r.Key).Distinct().Count();
var usedColumns = references.Where(r => r.Kind == "Column").Select(r => r.Key).Distinct().Count();
if (catalog != null)
{
    summary.AppendLine("Model definition: " + catalog.Name + " (" + catalog.Tables.Count + " tables)");
    summary.AppendLine("Model measures used: " + usedMeasures + " of " + catalog.Measures.Count);
    summary.AppendLine("Model columns used:  " + usedColumns + " of " + catalog.Columns.Count);
}
else
{
    summary.AppendLine("Verified model measures used: " + usedMeasures + ", columns used: " + usedColumns + ".");
    if (references.Any(r => r.Kind == unverifiedBare || r.Kind == unverifiedColumn))
        summary.AppendLine("Rows marked (unverified) come from apps whose model could not be read.");
}
summary.AppendLine("Bare [Name] references named as a result column in the same query (\"Week\") are listed as Query alias.");
Info(summary.ToString());

Output(dt);
