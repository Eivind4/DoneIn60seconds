// ============================================================================
// 14 - Detect and repair broken PBIR references.cs
//
// Scans one or more PBIR reports for table, measure, and column references
// that do not exist in the semantic model currently open in Tabular Editor.
//
// Replacement behavior:
//   - Missing table: select a replacement table. Only the table/entity name is
//     changed; field/property names are preserved for separate validation.
//   - Missing measure: select an existing measure.
//   - Missing column: select an existing column.
//   - Repeated references are grouped and repaired across all occurrences.
//   - Objects are reviewed one at a time. Each can be replaced or skipped.
//
// The macro updates structural field references, queryRef, nativeQueryRef when
// it still equals the old field name, selector.metadata, filter.From aliases,
// and nested advanced-filter Measure/Column expressions.
//
// Workflow:
//   1. Pick a folder. The macro finds every definition.pbir below it.
//   2. Select which discovered reports to include.
//
// IMPORTANT:
//   - PBIR edits are not undoable through Tabular Editor's undo stack. If the
//     reports are under source control, commit/sync first if required.
//   - The open Tabular Editor model is the source of truth. For a live
//     connection, ensure it represents the semantic model deployed to the
//     service; otherwise valid service fields may be reported as missing.
//   - After the update, apply external changes using Desktop Bridge, or close
//     and re-open the PBIR file in Power BI Desktop to see the changes.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

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

class BrokenReference
{
    public string Kind;
    public string Table;
    public string Field;
    public int Occurrences;
    public HashSet<string> Files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    public string Key
    {
        get { return Kind == "Table" ? "Table|" + Table : Kind + "|" + Table + "|" + Field; }
    }
}

class ReferenceReplacement
{
    public string Kind;
    public string OldTable;
    public string OldField;
    public string NewTable;
    public string NewField;
    public string Key
    {
        get { return Kind == "Table" ? "Table|" + OldTable : Kind + "|" + OldTable + "|" + OldField; }
    }
}

class PbirDocument
{
    public string DefinitionFolder;
    public string File;
    public JObject Root;
}

class FieldChoice
{
    public string Kind;
    public string Table;
    public string Name;
    public string DisplayFolder;
    public string Label
    {
        get { return Table + " > " + (string.IsNullOrWhiteSpace(DisplayFolder) ? "" : DisplayFolder + " > ") + Name; }
    }
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

// Read-only grid preview with a heading and Continue/Cancel buttons.
// Used instead of Output(), whose dialog only offers a "Close" button.
Func<string, string, System.Data.DataTable, string, bool> ShowGrid = (title, heading, table, continueText) =>
{
    using (var form = new Form())
    {
        form.Text = title;
        form.Size = BigDialogSize();
        form.MinimumSize = new Size(760, 480);
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.Sizable;
        form.MaximizeBox = true;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var label = new Label {
            Text = heading, Dock = DockStyle.Fill, AutoSize = true, MaximumSize = new Size(1100, 0),
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
// Step 1: select root folder and PBIR reports
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
        Text = "Selected reports are checked against the semantic model currently open in Tabular Editor. For live-connected reports, make sure the open model matches the one deployed to the service.",
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

var reportFolders = reportChoices
    .Where(r => r.Selected)
    .Select(r => r.ReportFolder)
    .Distinct(StringComparer.OrdinalIgnoreCase)
    .ToList();

// ---------------------------------------------------------------------------
// Step 2: load PBIR JSON documents and scan structural field references
// ---------------------------------------------------------------------------
var modelTables = new HashSet<string>(
    Model.Tables.Select(t => t.Name),
    StringComparer.OrdinalIgnoreCase);
var modelMeasures = new HashSet<string>(
    Model.AllMeasures.Select(m => m.Table.Name + "|" + m.Name),
    StringComparer.OrdinalIgnoreCase);
var modelColumns = new HashSet<string>(
    Model.AllColumns.Select(c => c.Table.Name + "|" + c.Name),
    StringComparer.OrdinalIgnoreCase);

var documents = new List<PbirDocument>();
var brokenByKey = new Dictionary<string, BrokenReference>(StringComparer.OrdinalIgnoreCase);

Action<string, string, string, string> recordBroken = (kind, table, field, file) =>
{
    if (string.IsNullOrEmpty(table)) return;
    var key = kind == "Table" ? "Table|" + table : kind + "|" + table + "|" + field;
    BrokenReference broken;
    if (!brokenByKey.TryGetValue(key, out broken))
    {
        broken = new BrokenReference {
            Kind = kind,
            Table = table,
            Field = field,
            Occurrences = 0
        };
        brokenByKey[key] = broken;
    }
    broken.Occurrences++;
    broken.Files.Add(file);
};

Action<JToken, string, Dictionary<string, string>, string> scan = null;
scan = (token, parentKey, aliasMap, file) =>
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
                if (!modelTables.Contains(table))
                    recordBroken("Table", table, null, file);
                else if (parentKey == "Measure" && !modelMeasures.Contains(table + "|" + field))
                    recordBroken("Measure", table, field, file);
                else if (parentKey == "Column" && !modelColumns.Contains(table + "|" + field))
                    recordBroken("Column", table, field, file);
            }
        }

        foreach (var property in obj.Properties())
            scan(property.Value, property.Name, localAliases, file);
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array) scan(item, parentKey, aliasMap, file);
    }
};

foreach (var reportFolder in reportFolders)
{
    var definitionFolder = Path.Combine(reportFolder, "definition");
    if (!Directory.Exists(definitionFolder))
    {
        Warning("Skipping selected report because its definition folder was not found: " + reportFolder);
        continue;
    }

    var candidateFiles = new List<string>();
    var reportJson = Path.Combine(definitionFolder, "report.json");
    if (File.Exists(reportJson)) candidateFiles.Add(reportJson);
    candidateFiles.AddRange(Directory.GetFiles(definitionFolder, "page.json", SearchOption.AllDirectories));
    candidateFiles.AddRange(Directory.GetFiles(definitionFolder, "visual.json", SearchOption.AllDirectories));

    foreach (var file in candidateFiles.Distinct(StringComparer.OrdinalIgnoreCase))
    {
        JObject root;
        try
        {
            root = JObject.Parse(File.ReadAllText(file));
        }
        catch (Exception ex)
        {
            Warning("Skipping unreadable PBIR JSON file:\n" + file + "\n\n" + ex.Message);
            continue;
        }

        documents.Add(new PbirDocument {
            DefinitionFolder = definitionFolder,
            File = file,
            Root = root
        });
        scan(root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), file);
    }
}

var brokenReferences = brokenByKey.Values
    .OrderBy(b => b.Kind == "Table" ? 0 : b.Kind == "Measure" ? 1 : 2)
    .ThenBy(b => b.Table)
    .ThenBy(b => b.Field)
    .ToList();

if (brokenReferences.Count == 0)
{
    Info("No broken table, measure, or column references were found in the selected PBIR reports.");
    return;
}

var findings = new System.Data.DataTable();
findings.Columns.Add("Type");
findings.Columns.Add("Missing table");
findings.Columns.Add("Missing field");
findings.Columns.Add("Occurrences", typeof(int));
findings.Columns.Add("Files", typeof(int));
foreach (var broken in brokenReferences)
    findings.Rows.Add(broken.Kind, broken.Table, broken.Field ?? "", broken.Occurrences, broken.Files.Count);
if (!ShowGrid(
        "Broken PBIR references",
        brokenReferences.Count + " broken reference(s) found. Select Continue to review and choose replacements one at a time.",
        findings,
        "Continue"))
{
    Info("Cancelled - no files were changed.");
    return;
}

// ---------------------------------------------------------------------------
// Shared picker layout: bold heading (row 0), search label (row 1), search box
// (row 2, added by caller), list/tree fill row (row 3), button row (row 4).
// ---------------------------------------------------------------------------
Func<Form, string, string, TableLayoutPanel> CreatePickerLayout = (form, heading, searchLabel) =>
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    form.Text = heading;
    form.Size = new Size((int)(wa.Width * 0.5), (int)(wa.Height * 0.75));
    form.MinimumSize = new Size(560, 440);
    form.StartPosition = FormStartPosition.CenterScreen;
    form.FormBorderStyle = FormBorderStyle.Sizable;
    form.MaximizeBox = true;

    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(12) };
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.Controls.Add(new Label {
        Text = heading + "\nSelect a replacement, or Skip to leave these references unchanged.",
        AutoSize = true, MaximumSize = new Size(900, 0),
        Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 10)
    }, 0, 0);
    layout.Controls.Add(new Label { Text = searchLabel, AutoSize = true, Font = new Font("Segoe UI", 9) }, 0, 1);
    form.Controls.Add(layout);
    return layout;
};

// ---------------------------------------------------------------------------
// Reusable single-select table picker
// ---------------------------------------------------------------------------
Func<string, string> selectTable = title =>
{
    string selected = null;
    using (var form = new Form())
    {
        var layout = CreatePickerLayout(form, title, "Search tables:");
        var font = new Font("Segoe UI", 9);
        var search = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var list = new ListBox { Dock = DockStyle.Fill, Font = font, IntegralHeight = false };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var okButton = new Button { Text = "Replace", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var skipButton = new Button { Text = "Skip", AutoSize = true, Font = font, DialogResult = DialogResult.Ignore, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { okButton, skipButton, cancelButton });
        list.DoubleClick += (s, e) => okButton.PerformClick();

        Action rebuild = () =>
        {
            list.Items.Clear();
            var text = (search.Text ?? "").Trim();
            foreach (var table in Model.Tables.Select(t => t.Name).OrderBy(n => n))
                if (string.IsNullOrEmpty(text) || table.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                    list.Items.Add(table);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        };

        search.TextChanged += (s, e) => rebuild();
        okButton.Click += (s, e) =>
        {
            if (list.SelectedItem == null)
            {
                MessageBox.Show("Select a replacement table.", "No table selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            selected = list.SelectedItem.ToString();
            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        form.AcceptButton = okButton;
        form.CancelButton = cancelButton;
        layout.Controls.Add(search, 0, 2);
        layout.Controls.Add(list, 0, 3);
        layout.Controls.Add(buttons, 0, 4);
        form.Shown += (s, e) => search.Focus();
        rebuild();

        var result = form.ShowDialog();
        if (result == DialogResult.Cancel) return null;
        if (result == DialogResult.Ignore) return "";
    }
    return selected;
};

// ---------------------------------------------------------------------------
// Reusable same-type field picker: Table > Display folder > Field
// ---------------------------------------------------------------------------
Func<string, string, FieldChoice> selectField = (kind, title) =>
{
    FieldChoice selected = null;
    var choices = kind == "Measure"
        ? Model.AllMeasures.Select(m => new FieldChoice {
            Kind = "Measure", Table = m.Table.Name, Name = m.Name, DisplayFolder = m.DisplayFolder ?? ""
        }).ToList()
        : Model.AllColumns.Select(c => new FieldChoice {
            Kind = "Column", Table = c.Table.Name, Name = c.Name, DisplayFolder = c.DisplayFolder ?? ""
        }).ToList();

    using (var form = new Form())
    {
        var layout = CreatePickerLayout(form, title, "Search table, display folder, or " + kind.ToLower() + ":");
        var font = new Font("Segoe UI", 9);
        var search = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var tree = new TreeView { Dock = DockStyle.Fill, Font = font, HideSelection = false };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var okButton = new Button { Text = "Replace", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var skipButton = new Button { Text = "Skip", AutoSize = true, Font = font, DialogResult = DialogResult.Ignore, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { okButton, skipButton, cancelButton });

        Action rebuild = () =>
        {
            tree.BeginUpdate();
            tree.Nodes.Clear();
            var text = (search.Text ?? "").Trim();
            var filtered = choices.Where(c =>
                string.IsNullOrEmpty(text)
                || c.Table.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                || c.DisplayFolder.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                || c.Name.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
            ).OrderBy(c => c.Table).ThenBy(c => c.DisplayFolder).ThenBy(c => c.Name).ToList();

            var nodeLookup = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var choice in filtered)
            {
                var parts = new List<string>() { choice.Table };
                if (!string.IsNullOrWhiteSpace(choice.DisplayFolder))
                    parts.AddRange(choice.DisplayFolder.Split('\\').Where(p => !string.IsNullOrWhiteSpace(p)));

                TreeNodeCollection level = tree.Nodes;
                string path = "";
                foreach (var part in parts)
                {
                    path = string.IsNullOrEmpty(path) ? part : path + "\\" + part;
                    TreeNode node;
                    if (!nodeLookup.TryGetValue(path, out node))
                    {
                        node = new TreeNode(part);
                        level.Add(node);
                        nodeLookup[path] = node;
                    }
                    level = node.Nodes;
                }
                level.Add(new TreeNode(choice.Name) { Tag = choice });
            }
            tree.ExpandAll();
            tree.EndUpdate();
        };

        search.TextChanged += (s, e) => rebuild();
        tree.NodeMouseDoubleClick += (s, e) =>
        {
            var choice = e.Node.Tag as FieldChoice;
            if (choice == null) return;
            selected = choice;
            form.DialogResult = DialogResult.OK;
            form.Close();
        };
        okButton.Click += (s, e) =>
        {
            var choice = tree.SelectedNode == null ? null : tree.SelectedNode.Tag as FieldChoice;
            if (choice == null)
            {
                MessageBox.Show("Select a replacement " + kind.ToLower() + ".", "No field selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            selected = choice;
            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        form.AcceptButton = okButton;
        form.CancelButton = cancelButton;
        layout.Controls.Add(search, 0, 2);
        layout.Controls.Add(tree, 0, 3);
        layout.Controls.Add(buttons, 0, 4);
        form.Shown += (s, e) => search.Focus();
        rebuild();

        var result = form.ShowDialog();
        if (result == DialogResult.Cancel) return null;
        if (result == DialogResult.Ignore) return new FieldChoice { Kind = "Skip" };
    }
    return selected;
};

// ---------------------------------------------------------------------------
// Step 3: review each grouped broken object and choose replacement or Skip
// ---------------------------------------------------------------------------
var replacements = new List<ReferenceReplacement>();

foreach (var broken in brokenReferences)
{
    if (broken.Kind == "Table")
    {
        var replacementTable = selectTable(
            "Missing table: " + broken.Table + " (" + broken.Occurrences + " reference(s))");
        if (replacementTable == null) return;
        if (replacementTable == "") continue;

        replacements.Add(new ReferenceReplacement {
            Kind = "Table",
            OldTable = broken.Table,
            NewTable = replacementTable
        });
    }
    else
    {
        var replacementField = selectField(
            broken.Kind,
            "Missing " + broken.Kind.ToLower() + ": " + broken.Table + "." + broken.Field +
            " (" + broken.Occurrences + " reference(s))");
        if (replacementField == null) return;
        if (replacementField.Kind == "Skip") continue;

        replacements.Add(new ReferenceReplacement {
            Kind = broken.Kind,
            OldTable = broken.Table,
            OldField = broken.Field,
            NewTable = replacementField.Table,
            NewField = replacementField.Name
        });
    }
}

if (replacements.Count == 0)
{
    Info("No replacements were selected. No files were changed.");
    return;
}

var tableMap = replacements
    .Where(r => r.Kind == "Table")
    .ToDictionary(r => r.OldTable, r => r.NewTable, StringComparer.OrdinalIgnoreCase);
var fieldMap = replacements
    .Where(r => r.Kind != "Table")
    .ToDictionary(r => r.Key, r => r, StringComparer.OrdinalIgnoreCase);

// ---------------------------------------------------------------------------
// Step 4: apply replacements in memory, including filter alias expressions
// ---------------------------------------------------------------------------
Func<JToken, string, Dictionary<string, string>, ReferenceReplacement, int> mutate = null;
mutate = (token, parentKey, aliasMap, forcedFilterReplacement) =>
{
    int changes = 0;
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
                if (string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(entity)) continue;

                string replacementTable;
                if (tableMap.TryGetValue(entity, out replacementTable))
                {
                    from["Entity"] = replacementTable;
                    entity = replacementTable;
                    changes++;
                }
                else if (forcedFilterReplacement != null
                    && string.Equals(entity, forcedFilterReplacement.OldTable, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(entity, forcedFilterReplacement.NewTable, StringComparison.OrdinalIgnoreCase))
                {
                    from["Entity"] = forcedFilterReplacement.NewTable;
                    entity = forcedFilterReplacement.NewTable;
                    changes++;
                }
                localAliases[alias] = entity;
            }
        }

        ReferenceReplacement localForced = forcedFilterReplacement;
        var outerMeasure = obj["field"]?["Measure"];
        var outerColumn = obj["field"]?["Column"];
        var outerField = outerMeasure ?? outerColumn;
        if (outerField != null && obj["filter"] != null)
        {
            var kind = outerMeasure != null ? "Measure" : "Column";
            var table = (string)outerField["Expression"]?["SourceRef"]?["Entity"];
            var field = (string)outerField["Property"];
            ReferenceReplacement hit;
            if (table != null && field != null
                && fieldMap.TryGetValue(kind + "|" + table + "|" + field, out hit))
                localForced = hit;
        }

        if (parentKey == "Measure" || parentKey == "Column")
        {
            var sourceRef = obj["Expression"]?["SourceRef"];
            var entityToken = sourceRef?["Entity"];
            var table = (string)entityToken;
            var sourceAlias = (string)sourceRef?["Source"];
            if (table == null && sourceAlias != null)
                localAliases.TryGetValue(sourceAlias, out table);
            var field = (string)obj["Property"];

            if (table != null)
            {
                string replacementTable;
                if (entityToken != null && tableMap.TryGetValue(table, out replacementTable))
                {
                    sourceRef["Entity"] = replacementTable;
                    table = replacementTable;
                    changes++;
                }
            }

            if (table != null && field != null)
            {
                ReferenceReplacement hit;
                if (localForced != null
                    && localForced.Kind == parentKey
                    && (string.Equals(localForced.OldTable, table, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(localForced.NewTable, table, StringComparison.OrdinalIgnoreCase)))
                {
                    if (sourceAlias != null && fromArray != null)
                    {
                        foreach (var from in fromArray.OfType<JObject>())
                            if (string.Equals((string)from["Name"], sourceAlias, StringComparison.OrdinalIgnoreCase)
                                && !string.Equals((string)from["Entity"], localForced.NewTable, StringComparison.OrdinalIgnoreCase))
                            {
                                from["Entity"] = localForced.NewTable;
                                changes++;
                            }
                    }
                    if (entityToken != null && !string.Equals((string)entityToken, localForced.NewTable, StringComparison.OrdinalIgnoreCase))
                    {
                        sourceRef["Entity"] = localForced.NewTable;
                        changes++;
                    }
                    if (!string.Equals(field, localForced.NewField, StringComparison.Ordinal))
                    {
                        obj["Property"] = localForced.NewField;
                        changes++;
                    }
                }
                else if (fieldMap.TryGetValue(parentKey + "|" + table + "|" + field, out hit))
                {
                    if (entityToken != null && !string.Equals(table, hit.NewTable, StringComparison.OrdinalIgnoreCase))
                    {
                        sourceRef["Entity"] = hit.NewTable;
                        changes++;
                    }
                    obj["Property"] = hit.NewField;
                    changes++;
                }
            }
        }

        var queryRef = obj["queryRef"];
        if (queryRef != null && queryRef.Type == JTokenType.String)
        {
            var value = (string)queryRef;
            var fieldHit = fieldMap.Values.FirstOrDefault(r =>
                string.Equals(value, r.OldTable + "." + r.OldField, StringComparison.OrdinalIgnoreCase));
            if (fieldHit != null)
            {
                obj["queryRef"] = fieldHit.NewTable + "." + fieldHit.NewField;
                changes++;
                var native = obj["nativeQueryRef"];
                if (native != null && native.Type == JTokenType.String
                    && string.Equals((string)native, fieldHit.OldField, StringComparison.Ordinal))
                {
                    obj["nativeQueryRef"] = fieldHit.NewField;
                    changes++;
                }
            }
            else
            {
                foreach (var tableReplacement in tableMap)
                {
                    var prefix = tableReplacement.Key + ".";
                    if (value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        obj["queryRef"] = tableReplacement.Value + value.Substring(tableReplacement.Key.Length);
                        changes++;
                        break;
                    }
                }
            }
        }

        var metadata = obj["metadata"];
        if (metadata != null && metadata.Type == JTokenType.String)
        {
            var value = (string)metadata;
            var fieldHit = fieldMap.Values.FirstOrDefault(r =>
                string.Equals(value, r.OldTable + "." + r.OldField, StringComparison.OrdinalIgnoreCase));
            if (fieldHit != null)
            {
                obj["metadata"] = fieldHit.NewTable + "." + fieldHit.NewField;
                changes++;
            }
            else
            {
                foreach (var tableReplacement in tableMap)
                {
                    var prefix = tableReplacement.Key + ".";
                    if (value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        obj["metadata"] = tableReplacement.Value + value.Substring(tableReplacement.Key.Length);
                        changes++;
                        break;
                    }
                }
            }
        }

        foreach (var property in obj.Properties().ToList())
            changes += mutate(property.Value, property.Name, localAliases, localForced);
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array) changes += mutate(item, parentKey, aliasMap, forcedFilterReplacement);
    }
    return changes;
};

var filesToWrite = new List<Tuple<string, string, JObject, int>>();
foreach (var document in documents)
{
    int changes = mutate(
        document.Root,
        null,
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
        null);
    if (changes > 0)
        filesToWrite.Add(Tuple.Create(document.DefinitionFolder, document.File, document.Root, changes));
}

// ---------------------------------------------------------------------------
// Step 5: preview selected mappings and affected files
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
preview.Columns.Add("Scope");
preview.Columns.Add("Type");
preview.Columns.Add("Old reference");
preview.Columns.Add("New reference");
preview.Columns.Add("Action");

foreach (var replacement in replacements)
{
    var oldReference = replacement.Kind == "Table"
        ? replacement.OldTable
        : replacement.OldTable + "." + replacement.OldField;
    var newReference = replacement.Kind == "Table"
        ? replacement.NewTable
        : replacement.NewTable + "." + replacement.NewField;
    preview.Rows.Add("Mapping", replacement.Kind, oldReference, newReference, "Replace all grouped occurrences");
}

foreach (var file in filesToWrite)
{
    var relativePath = file.Item2.Substring(file.Item1.Length).TrimStart('\\', '/');
    preview.Rows.Add(file.Item1 + " | " + relativePath, "PBIR file", "", "", "Update " + file.Item4 + " reference value(s)");
}

if (filesToWrite.Count == 0)
{
    Output(preview);
    Info("The selected replacements did not produce any PBIR changes.");
    return;
}

if (!ShowGrid(
        "Preview broken-reference repair",
        "Review the mappings and affected files. Select Continue to confirm and write the changes.",
        preview,
        "Continue"))
{
    Info("Cancelled - no files were changed.");
    return;
}

var confirm = MessageBox.Show(
    "Apply " + replacements.Count + " grouped replacement(s) across " +
    filesToWrite.Count + " PBIR JSON file(s)?\n\n" +
    "PBIR edits cannot be undone in Tabular Editor. If the reports are under source control, commit/sync first if required.\n\n" +
    "Afterwards, apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes.",
    "Confirm broken-reference repair",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) return;

// ---------------------------------------------------------------------------
// Step 6: write files and summarize partial progress if an error occurs
// ---------------------------------------------------------------------------
int filesWritten = 0;
int referencesWritten = 0;

foreach (var file in filesToWrite)
{
    try
    {
        File.WriteAllText(file.Item2, file.Item3.ToString(Newtonsoft.Json.Formatting.Indented));
        filesWritten++;
        referencesWritten += file.Item4;
    }
    catch (Exception ex)
    {
        Error(string.Format(
            "The repair stopped while writing:\n{0}\n\n" +
            "Changes already made before the error:\n" +
            "- {1} PBIR reference value(s)\n" +
            "- {2} PBIR JSON file(s)\n\n" +
            "Error: {3}",
            file.Item2,
            referencesWritten,
            filesWritten,
            ex.Message));
    }
}

Info(string.Format(
    "Applied {0} grouped replacement(s) and updated {1} PBIR reference value(s) across {2} file(s).\n\n" +
    "Next: apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes.\n" +
    "If this is a live-connected report, ensure the referenced model is also deployed and current in the service.",
    replacements.Count,
    referencesWritten,
    filesWritten));
