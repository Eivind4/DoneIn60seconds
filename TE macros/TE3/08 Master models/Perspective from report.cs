// ============================================================================
// Perspective from report.cs
//
// Tabular Editor 3 macro that reads one or more PBIR reports and creates a
// perspective in the open model with exactly the fields each report uses.
// "13 - Master model perspectives" can then turn the perspective into a
// smaller model for that report.
//
// Workflow:
//   1. Pick a folder. The macro finds every definition.pbir below it.
//   2. A grid lists the reports. Tick the ones to use and edit the perspective
//      name (default: the report name). Give two reports the same name to
//      combine them in one perspective.
//   3. Fields are collected from report.json, page.json, visual.json and
//      bookmarks: measures, columns and hierarchies, wherever they are used
//      (fields, filters, sort, conditional formatting, SourceRef aliases).
//      Report-level measures (reportExtensions.json) are not model objects;
//      the model measures and columns their DAX references are included.
//   4. Option (default on): also include every measure the used measures
//      depend on (directly or through other measures and DAX UDFs), so
//      "13 - Master model perspectives" accepts the perspective as it is.
//   5. Existing perspective: Replace (membership becomes the report's fields)
//      or Add (the report's fields are added to what is already there).
//   6. A preview lists every field per perspective, including fields the
//      report uses that are not in the model (renamed or deleted objects).
//      Nothing is changed before you confirm.
//
// Notes:
//   - Only the model is changed (perspective membership), and Ctrl+Z works
//     until you save. The PBIR files are only read.
//   - Columns that measures need are not added: "13 - Master model
//     perspectives" keeps them, hidden, in the smaller model.
//   - Visual calculations are not parsed. DAX text in report-level measures
//     is matched with [Measure] and 'Table'[Column] patterns only.
//   - Reports must be saved in PBIR format (not PBIX or PBIR-Legacy).
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

class ReportRef
{
    public string Folder;
    public string Name;
    public HashSet<string> Measures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);     // Table|Name
    public HashSet<string> Columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);      // Table|Name
    public HashSet<string> Fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);       // Table|Name, measure or column (selector metadata)
    public HashSet<string> Hierarchies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);  // Table|Hierarchy
    public HashSet<string> Levels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);       // Table|Hierarchy|Level
    public HashSet<string> ExtensionDax = new HashSet<string>();                                 // expressions of report-level measures
}

class PlanRow
{
    public string Perspective;
    public string Kind;
    public string Table;
    public string Name;
    public string Reason;
    public object Obj;          // Measure, Column or Hierarchy (null = not in the model)
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
string CleanReportName(string folder)
{
    var name = Path.GetFileName(folder);
    return name.EndsWith(".Report", StringComparison.OrdinalIgnoreCase) ? name.Substring(0, name.Length - ".Report".Length) : name;
}

Table FindTable(string name) { return Model.Tables.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)); }

Measure FindMeasure(string table, string name)
{
    var t = FindTable(table);
    var m = t == null ? null : t.Measures.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
    // Measure names are unique in the model, so a measure moved to another table is still found
    return m ?? Model.AllMeasures.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
}

Column FindColumn(string table, string name)
{
    var t = FindTable(table);
    return t == null ? null : t.Columns.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
}

Hierarchy FindHierarchy(string table, string name)
{
    var t = FindTable(table);
    return t == null ? null : t.Hierarchies.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
}

// DependsOn in the TE object model: enumerates keys (dictionary) or objects
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

// Structural references in PBIR JSON. SourceRef points to a table by Entity, or to an
// alias (Source) declared in a From array higher up. Report-level measures have Schema "extension".
void ScanToken(JToken token, string parentKey, Dictionary<string, string> aliases, HashSet<string> extensionAliases, ReportRef report)
{
    var obj = token as JObject;
    if (obj == null)
    {
        var array = token as JArray;
        if (array != null) foreach (var item in array) ScanToken(item, parentKey, aliases, extensionAliases, report);
        return;
    }

    var localAliases = aliases;
    var localExtension = extensionAliases;
    var from = obj["From"] as JArray;
    if (from != null)
    {
        localAliases = new Dictionary<string, string>(aliases, StringComparer.OrdinalIgnoreCase);
        localExtension = new HashSet<string>(extensionAliases, StringComparer.OrdinalIgnoreCase);
        foreach (var f in from.OfType<JObject>())
        {
            var alias = (string)f["Name"];
            var entity = (string)f["Entity"];
            if (string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(entity)) continue;
            localAliases[alias] = entity;
            if ((string)f["Schema"] == "extension") localExtension.Add(alias); else localExtension.Remove(alias);
        }
    }

    Func<JToken, string> tableOf = expression =>
    {
        var sourceRef = expression == null ? null : expression["SourceRef"];
        if (sourceRef == null) return null;
        if ((string)sourceRef["Schema"] == "extension") return null;
        var entity = (string)sourceRef["Entity"];
        if (entity != null) return entity;
        var source = (string)sourceRef["Source"];
        string aliased;
        if (source == null || localExtension.Contains(source) || !localAliases.TryGetValue(source, out aliased)) return null;
        return aliased;
    };

    if (parentKey == "Measure" || parentKey == "Column")
    {
        var table = tableOf(obj["Expression"]);
        var field = (string)obj["Property"];
        if (table != null && !string.IsNullOrEmpty(field))
            (parentKey == "Measure" ? report.Measures : report.Columns).Add(table + "|" + field);
    }
    else if (parentKey == "Hierarchy" && obj["Hierarchy"] is JValue)
    {
        var table = tableOf(obj["Expression"]);
        if (table != null) report.Hierarchies.Add(table + "|" + (string)obj["Hierarchy"]);
    }
    else if (parentKey == "HierarchyLevel" && obj["Level"] is JValue)
    {
        var hierarchy = obj["Expression"] == null ? null : obj["Expression"]["Hierarchy"] as JObject;
        var table = hierarchy == null ? null : tableOf(hierarchy["Expression"]);
        if (table != null) report.Levels.Add(table + "|" + (string)hierarchy["Hierarchy"] + "|" + (string)obj["Level"]);
    }

    // Formatting entries scoped to one field: selector.metadata = "Table.Field"
    var metadata = obj["selector"] == null ? null : (string)obj["selector"]["metadata"];
    if (!string.IsNullOrWhiteSpace(metadata) && metadata.IndexOf('(') < 0)
    {
        var dot = metadata.IndexOf('.');
        if (dot > 0 && dot < metadata.Length - 1) report.Fields.Add(metadata.Substring(0, dot) + "|" + metadata.Substring(dot + 1));
    }

    foreach (var property in obj.Properties()) ScanToken(property.Value, property.Name, localAliases, localExtension, report);
}

void ScanReport(ReportRef report)
{
    var definition = Path.Combine(report.Folder, "definition");
    if (!Directory.Exists(definition)) return;
    foreach (var file in Directory.GetFiles(definition, "*.json", SearchOption.AllDirectories))
    {
        var fileName = Path.GetFileName(file);
        try
        {
            var root = JObject.Parse(File.ReadAllText(file));
            if (string.Equals(fileName, "reportExtensions.json", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var entity in (root["entities"] as JArray ?? new JArray()).OfType<JObject>())
                    foreach (var m in (entity["measures"] as JArray ?? new JArray()).OfType<JObject>())
                    {
                        var expression = (string)m["expression"];
                        if (!string.IsNullOrWhiteSpace(expression)) report.ExtensionDax.Add(expression);
                        var refs = m["references"] == null ? null : m["references"]["measures"] as JArray;
                        if (refs != null)
                            foreach (var r in refs.OfType<JObject>().Where(r => (string)r["schema"] != "extension"))
                                report.Measures.Add((string)r["entity"] + "|" + (string)r["name"]);
                    }
                continue;
            }
            ScanToken(root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), new HashSet<string>(StringComparer.OrdinalIgnoreCase), report);
        }
        catch (Exception ex) { Warning("Skipping unreadable PBIR file:\n" + file + "\n\n" + ex.Message); }
    }
}

// ---------------------------------------------------------------------------
// Step 1: folder and reports
// ---------------------------------------------------------------------------
string rootFolder;
using (var fbd = new FolderBrowserDialog())
{
    fbd.Description = "Select a folder with one or more PBIR reports (.Report folders with definition.pbir).";
    if (fbd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    rootFolder = fbd.SelectedPath;
}

var reports = Directory.GetFiles(rootFolder, "definition.pbir", SearchOption.AllDirectories)
    .Select(f => new ReportRef { Folder = Path.GetDirectoryName(f), Name = CleanReportName(Path.GetDirectoryName(f)) })
    .OrderBy(r => r.Name).ToList();
if (reports.Count == 0) { Info("No definition.pbir files were found below:\n" + rootFolder); return; }

// ---------------------------------------------------------------------------
// Step 2: reports -> perspective names, options
// ---------------------------------------------------------------------------
var chosen = new List<Tuple<ReportRef, string>>();   // report, perspective name
bool includeDependencies = true;
bool replaceExisting = true;
using (var form = new Form())
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    form.Text = "Perspective from report";
    form.Size = new Size((int)(wa.Width * 0.6), (int)(wa.Height * 0.6));
    form.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);
    form.Font = font;

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(12) };
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var heading = new Label {
        Text = "Tick the reports and edit the perspective names. Give two reports the same name to combine them in one perspective.",
        AutoSize = true, MaximumSize = new Size(form.Width - 60, 0), Margin = new Padding(0, 0, 0, 8)
    };
    var grid = new DataGridView {
        Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.CellSelect, BackgroundColor = SystemColors.Window
    };
    grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Include", HeaderText = "Include", Width = 60 });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Report", HeaderText = "Report", Width = 220, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Perspective", HeaderText = "Perspective name (edit)", Width = 220 });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "Status", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
    Action<DataGridViewRow> refresh = row =>
    {
        var name = Convert.ToString(row.Cells["Perspective"].Value).Trim();
        row.Cells["Status"].Value = name.Length == 0 ? "Enter a name"
            : Model.Perspectives.Any(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) ? "Perspective exists" : "New perspective";
    };
    foreach (var r in reports)
    {
        var index = grid.Rows.Add(reports.Count == 1, r.Name, r.Name, "");
        grid.Rows[index].Tag = r;
        refresh(grid.Rows[index]);
    }
    grid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0 && grid.Columns[e.ColumnIndex].Name == "Perspective") refresh(grid.Rows[e.RowIndex]); };

    var dependencies = new CheckBox {
        Text = "Also include the measures the used measures depend on (required when building a smaller model per perspective)",
        Checked = true, AutoSize = true, Margin = new Padding(0, 8, 0, 0)
    };
    var replace = new RadioButton { Text = "Existing perspective: replace its fields with the report's fields", Checked = true, AutoSize = true };
    var add = new RadioButton { Text = "Existing perspective: add the report's fields to it", AutoSize = true };
    var radios = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(0, 4, 0, 0) };
    radios.Controls.AddRange(new Control[] { replace, add });

    var buttons = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
    var ok = new Button { Text = "Continue", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
    var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { ok, cancel });
    ok.Click += (s, e) =>
    {
        grid.EndEdit();
        chosen.Clear();
        foreach (DataGridViewRow row in grid.Rows)
        {
            if (!Convert.ToBoolean(row.Cells["Include"].Value)) continue;
            var name = Convert.ToString(row.Cells["Perspective"].Value).Trim();
            if (name.Length == 0) { MessageBox.Show("Enter a perspective name for " + row.Cells["Report"].Value + ".", "Perspective name"); return; }
            chosen.Add(Tuple.Create((ReportRef)row.Tag, name));
        }
        if (chosen.Count == 0) { MessageBox.Show("Tick at least one report.", "Nothing selected"); return; }
        includeDependencies = dependencies.Checked;
        replaceExisting = replace.Checked;
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    layout.Controls.Add(heading, 0, 0);
    layout.Controls.Add(grid, 0, 1);
    layout.Controls.Add(dependencies, 0, 2);
    layout.Controls.Add(radios, 0, 3);
    layout.Controls.Add(buttons, 0, 5);
    form.Controls.Add(layout);
    form.CancelButton = cancel;
    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 3: scan the reports and resolve the fields in the model
// ---------------------------------------------------------------------------
foreach (var report in chosen.Select(c => c.Item1).Distinct()) ScanReport(report);

var columnInDax = new Regex(@"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))\[([^\]]+)\]");
var measureInDax = new Regex(@"(?<![\w'\]])\[([^\]]+)\]");

var plan = new List<PlanRow>();
foreach (var group in chosen.GroupBy(c => c.Item2, StringComparer.OrdinalIgnoreCase))
{
    var perspective = group.First().Item2;
    var rows = new Dictionary<object, PlanRow>();
    var missing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    Action<object, string, string, string, string> addRow = (obj, kind, table, name, reason) =>
    {
        if (obj == null)
        {
            if (missing.Add(kind + "|" + table + "|" + name))
                plan.Add(new PlanRow { Perspective = perspective, Kind = kind, Table = table, Name = name, Reason = "NOT IN MODEL - " + reason });
            return;
        }
        if (rows.ContainsKey(obj)) return;
        var row = new PlanRow { Perspective = perspective, Kind = kind, Table = table, Name = name, Reason = reason, Obj = obj };
        rows[obj] = row;
        plan.Add(row);
    };

    foreach (var report in group.Select(c => c.Item1))
    {
        var used = "used in " + report.Name;
        foreach (var key in report.Measures)
        {
            var p = key.Split('|');
            var m = FindMeasure(p[0], p[1]);
            addRow(m, "Measure", m != null ? m.Table.Name : p[0], p[1], used);
        }
        foreach (var key in report.Columns)
        {
            var p = key.Split('|');
            addRow(FindColumn(p[0], p[1]), "Column", p[0], p[1], used);
        }
        foreach (var key in report.Fields)
        {
            var p = key.Split('|');
            var m = FindMeasure(p[0], p[1]);
            if (m != null) { addRow(m, "Measure", m.Table.Name, m.Name, used + " (formatting)"); continue; }
            var c = FindColumn(p[0], p[1]);
            if (c != null) addRow(c, "Column", p[0], p[1], used + " (formatting)");
        }
        foreach (var key in report.Hierarchies)
        {
            var p = key.Split('|');
            addRow(FindHierarchy(p[0], p[1]), "Hierarchy", p[0], p[1], used);
        }
        foreach (var key in report.Levels)
        {
            var p = key.Split('|');
            var h = FindHierarchy(p[0], p[1]);
            var level = h == null ? null : h.Levels.FirstOrDefault(l => string.Equals(l.Name, p[2], StringComparison.OrdinalIgnoreCase));
            addRow(h, "Hierarchy", p[0], p[1], used + " (level " + p[2] + ")");
            if (level != null) addRow(level.Column, "Column", p[0], level.Column.Name, used + " (level of " + p[1] + ")");
        }
        // Report-level measures: model measures and columns named in their DAX (others are report-level measures)
        foreach (var dax in report.ExtensionDax)
        {
            foreach (Match match in columnInDax.Matches(dax))
            {
                var table = match.Groups[1].Success ? match.Groups[1].Value.Replace("''", "'") : match.Groups[2].Value;
                var c = FindColumn(table, match.Groups[3].Value);
                if (c != null) addRow(c, "Column", c.Table.Name, c.Name, "used by a report-level measure in " + report.Name);
            }
            foreach (Match match in measureInDax.Matches(dax))
            {
                var m = Model.AllMeasures.FirstOrDefault(x => string.Equals(x.Name, match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
                if (m != null) addRow(m, "Measure", m.Table.Name, m.Name, "used by a report-level measure in " + report.Name);
            }
        }
    }

    if (includeDependencies)
    {
        var queue = new Queue<object>(rows.Keys.OfType<Measure>());
        var seen = new HashSet<object>(queue);
        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            var user = current as Measure;
            foreach (var dep in DependenciesOf(current))
            {
                var m = dep as Measure;
                if (m == null && !dep.GetType().Name.Contains("Function")) continue;
                if (!seen.Add(dep)) continue;
                if (m != null) addRow(m, "Measure", m.Table.Name, m.Name, "dependency" + (user != null ? " of " + user.DaxObjectFullName : " of a DAX UDF"));
                queue.Enqueue(dep);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Step 4: preview and confirm
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
foreach (var column in new[] { "Perspective", "Type", "Table", "Name", "Why" }) preview.Columns.Add(column);
foreach (var r in plan.OrderBy(r => r.Perspective).ThenBy(r => r.Obj == null ? 0 : 1).ThenBy(r => r.Kind).ThenBy(r => r.Table).ThenBy(r => r.Name))
    preview.Rows.Add(r.Perspective, r.Kind, r.Table, r.Name, r.Reason);
Output(preview);

var summaryLines = plan.GroupBy(r => r.Perspective).Select(g =>
{
    var exists = Model.Perspectives.Any(p => string.Equals(p.Name, g.Key, StringComparison.OrdinalIgnoreCase));
    var found = g.Where(r => r.Obj != null).ToList();
    var notFound = g.Count(r => r.Obj == null);
    return g.Key + (exists ? (replaceExisting ? " (replace)" : " (add)") : " (new)") + ": " +
        found.Count(r => r.Kind == "Measure") + " measures, " + found.Count(r => r.Kind == "Column") + " columns, " +
        found.Count(r => r.Kind == "Hierarchy") + " hierarchies" + (notFound > 0 ? ", " + notFound + " not in the model" : "");
}).ToList();
if (summaryLines.Count == 0) { Info("The selected reports use no fields from this model - no changes made."); return; }
if (MessageBox.Show("Update these perspectives?\n\n" + string.Join("\n", summaryLines) +
    "\n\nFields not in the model are listed in the preview and skipped. Ctrl+Z undoes the change until you save.",
    "Perspective from report", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
{ Info("Cancelled - no changes made."); return; }

// ---------------------------------------------------------------------------
// Step 5: apply
// ---------------------------------------------------------------------------
foreach (var group in plan.Where(r => r.Obj != null).GroupBy(r => r.Perspective, StringComparer.OrdinalIgnoreCase))
{
    var name = group.Key;
    var existing = Model.Perspectives.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
    if (existing == null) Model.AddPerspective(name);
    else
    {
        name = existing.Name;
        if (replaceExisting) foreach (var t in Model.Tables) t.InPerspective[name] = false;
    }
    foreach (var r in group)
    {
        if (r.Obj is Measure) ((Measure)r.Obj).InPerspective[name] = true;
        else if (r.Obj is Column) ((Column)r.Obj).InPerspective[name] = true;
        else if (r.Obj is Hierarchy) ((Hierarchy)r.Obj).InPerspective[name] = true;
    }
}

Info("Done.\n\n" + string.Join("\n", summaryLines) +
    "\n\nSave the model to keep the perspectives. A smaller model per perspective can then be built from the master model.");
