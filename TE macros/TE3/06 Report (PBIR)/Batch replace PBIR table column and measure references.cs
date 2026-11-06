// ============================================================================
// 15 - Batch replace PBIR table column and measure references.cs
//
// General TOM + PBIR reference replacement macro for tables, columns, and measures.
// It inventories references found in one or more reports, flags whether each
// reference exists in the semantic model open in Tabular Editor, and lets you
// enter new names for grouped references in one editable grid.
//
// Which references:
//   - If tables, measures, or columns are selected in Tabular Editor, ONLY
//     those objects are processed (also when a selected object is not used in
//     the reports, so it can still be renamed in the model).
//   - Otherwise, pick references from a tree:
//       Tables > Table
//       Measures / Columns > Table > Field
//
// Scope:
//   - Always every selected report in full: all visual.json, page.json, and
//     report.json files. Renaming a model object while only updating some
//     visuals would break the visuals that were not updated.
//
// Replacement rules (per row in the rename grid):
//   - Source exists in the model and the new name is free: the model object is
//     renamed in TOM and PBIR is synchronized.
//   - New name already exists in the model: PBIR is repointed to that existing
//     object; the source model object is not renamed or overwritten. Measures
//     are repointed to the table that hosts the existing measure.
//   - Source missing from the model: PBIR-only repair.
//   - Columns and measures stay on their table (including a table renamed in
//     the same run). Repeated references are grouped and replaced together.
//   - projection.displayName is preserved.
//
// Updated PBIR surfaces include structural Measure/Column references,
// queryRef, nativeQueryRef when it still equals the old field name,
// selector.metadata, filter.From aliases, and nested advanced-filter
// Measure/Column expressions.
//
// IMPORTANT:
//   - PBIR edits are not undoable through Tabular Editor's undo stack (model
//     renames are, with Ctrl+Z). If the model/reports are under source
//     control, commit/sync first if required.
//   - After running, FIRST save the model changes in Tabular Editor and, if
//     required (e.g. live-connected reports), deploy/update the model in the
//     service. Updating only the report files is not sufficient.
//   - Then apply external changes using Desktop Bridge, or close and re-open
//     the PBIR file in Power BI Desktop to see the changes in the reports.
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

class GeneralReference
{
    public string Kind;
    public string Table;
    public string Field;
    public bool ExistsInModel;
    public int Occurrences;
    public HashSet<string> Files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    public string Key
    {
        get { return Kind == "Table" ? "Table|" + Table : Kind + "|" + Table + "|" + Field; }
    }
    public string Label
    {
        get { return Kind == "Table" ? Table : Table + "." + Field; }
    }
}

class GeneralReplacement
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

class GeneralModelAction
{
    public string ReplacementKey;
    public string Kind;
    public string Action;
    public string Detail;
    public Table TableObject;
    public Measure MeasureObject;
    public Column ColumnObject;
    public string NewName;
}

class GeneralDocument
{
    public string DefinitionFolder;
    public string File;
    public JObject Root;
    public string ReportName;
}

const string ActionRename = "Rename in TOM + sync PBIR";

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

Func<string, string, bool> SameName = (a, b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

// Standard resizable dialog with a TableLayoutPanel; rows are added by the caller.
Func<Form, string, int, TableLayoutPanel> CreateDialogLayout = (form, title, rowCount) =>
{
    form.Text = title;
    form.Size = BigDialogSize();
    form.MinimumSize = new Size(760, 480);
    form.StartPosition = FormStartPosition.CenterScreen;
    form.FormBorderStyle = FormBorderStyle.Sizable;
    form.MaximizeBox = true;
    var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = rowCount, ColumnCount = 1, Padding = new Padding(12) };
    form.Controls.Add(layout);
    return layout;
};

// Read-only grid preview with a heading and Continue/Cancel buttons.
// Used instead of Output(), whose dialog only offers a "Close" button.
Func<string, string, System.Data.DataTable, string, bool> ShowGrid = (title, heading, table, continueText) =>
{
    using (var form = new Form())
    {
        var layout = CreateDialogLayout(form, title, 3);
        var font = new Font("Segoe UI", 9);
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
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK;
    }
};

// ---------------------------------------------------------------------------
// Step 0: objects selected in Tabular Editor restrict the run to those objects
// ---------------------------------------------------------------------------
var modelSelectionKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
foreach (var t in Selected.Tables) modelSelectionKeys.Add("Table|" + t.Name);
foreach (var m in Selected.Measures) modelSelectionKeys.Add("Measure|" + m.Table.Name + "|" + m.Name);
foreach (var c in Selected.Columns) modelSelectionKeys.Add("Column|" + c.Table.Name + "|" + c.Name);
bool useModelSelection = modelSelectionKeys.Count > 0;

// ---------------------------------------------------------------------------
// Step 1: select root folder and PBIR reports
// ---------------------------------------------------------------------------
string rootFolder = null;
using (var fbd = new FolderBrowserDialog())
{
    fbd.Description = "Select a folder containing one or more PBIR .Report folders (definition.pbir files).";
    if (fbd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no folder selected. No changes made."); return; }
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
    Info("No definition.pbir files were found below:\n" + rootFolder + "\n\nNo changes made.");
    return;
}

using (var form = new Form())
{
    var layout = CreateDialogLayout(form, "Select PBIR reports to update", 4);
    var font = new Font("Segoe UI", 9);
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var warning = new Label {
        Text = (useModelSelection
                ? "Only the " + modelSelectionKeys.Count + " table(s)/measure(s)/column(s) selected in Tabular Editor will be processed. "
                : "No tables, measures, or columns are selected in Tabular Editor - you can pick references after scanning. ") +
            "Include every report that uses the model: renamed model objects break reports that are not updated.",
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
// Step 2: load report/page/visual JSON from every selected report
// ---------------------------------------------------------------------------
var documents = new List<GeneralDocument>();

foreach (var reportFolder in reportFolders)
{
    var definitionFolder = Path.Combine(reportFolder, "definition");
    if (!Directory.Exists(definitionFolder))
    {
        Warning("Skipping selected report because its definition folder was not found: " + reportFolder);
        continue;
    }
    var reportName = CleanReportName(reportFolder);

    var reportJson = Path.Combine(definitionFolder, "report.json");
    if (File.Exists(reportJson))
    {
        try
        {
            documents.Add(new GeneralDocument {
                DefinitionFolder = definitionFolder,
                File = reportJson,
                Root = JObject.Parse(File.ReadAllText(reportJson)),
                ReportName = reportName
            });
        }
        catch (Exception ex)
        {
            Warning("Skipping unreadable report.json:\n" + reportJson + "\n\n" + ex.Message);
        }
    }

    var pagesFolder = Path.Combine(definitionFolder, "pages");
    if (!Directory.Exists(pagesFolder)) continue;

    foreach (var pageFolder in Directory.GetDirectories(pagesFolder))
    {
        var pageFile = Path.Combine(pageFolder, "page.json");
        JObject pageRoot = null;
        if (File.Exists(pageFile))
        {
            try
            {
                pageRoot = JObject.Parse(File.ReadAllText(pageFile));
                documents.Add(new GeneralDocument {
                    DefinitionFolder = definitionFolder,
                    File = pageFile,
                    Root = pageRoot,
                    ReportName = reportName
                });
            }
            catch (Exception ex)
            {
                Warning("Skipping unreadable page.json:\n" + pageFile + "\n\n" + ex.Message);
            }
        }

        var visualsFolder = Path.Combine(pageFolder, "visuals");
        if (!Directory.Exists(visualsFolder)) continue;

        foreach (var file in Directory.GetFiles(visualsFolder, "visual.json", SearchOption.AllDirectories))
        {
            try
            {
                documents.Add(new GeneralDocument {
                    DefinitionFolder = definitionFolder,
                    File = file,
                    Root = JObject.Parse(File.ReadAllText(file)),
                    ReportName = reportName
                });
            }
            catch (Exception ex)
            {
                Warning("Skipping unreadable visual.json:\n" + file + "\n\n" + ex.Message);
            }
        }
    }
}

if (documents.Count == 0 && !useModelSelection)
{
    Info("No readable PBIR report, page, or visual JSON files were found. No changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 3: inventory all table, measure, and column references
// ---------------------------------------------------------------------------
var modelTables = new HashSet<string>(Model.Tables.Select(t => t.Name), StringComparer.OrdinalIgnoreCase);
var modelMeasures = new HashSet<string>(Model.AllMeasures.Select(m => m.Table.Name + "|" + m.Name), StringComparer.OrdinalIgnoreCase);
var modelColumns = new HashSet<string>(Model.AllColumns.Select(c => c.Table.Name + "|" + c.Name), StringComparer.OrdinalIgnoreCase);
var referencesByKey = new Dictionary<string, GeneralReference>(StringComparer.OrdinalIgnoreCase);

Func<string, string, string, GeneralReference> getOrAddReference = (kind, table, field) =>
{
    var key = kind == "Table" ? "Table|" + table : kind + "|" + table + "|" + field;
    GeneralReference reference;
    if (!referencesByKey.TryGetValue(key, out reference))
    {
        bool exists = kind == "Table"
            ? modelTables.Contains(table)
            : kind == "Measure"
                ? modelMeasures.Contains(table + "|" + field)
                : modelColumns.Contains(table + "|" + field);
        reference = new GeneralReference { Kind = kind, Table = table, Field = field, ExistsInModel = exists };
        referencesByKey[key] = reference;
    }
    return reference;
};

Action<string, string, string, string> recordReference = (kind, table, field, file) =>
{
    if (string.IsNullOrEmpty(table)) return;
    var reference = getOrAddReference(kind, table, field);
    reference.Occurrences++;
    reference.Files.Add(file);
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
                if (!string.IsNullOrEmpty(entity)) recordReference("Table", entity, null, file);
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
                recordReference("Table", table, null, file);
                recordReference(parentKey, table, field, file);
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

foreach (var document in documents)
    scan(document.Root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), document.File);

// Objects selected in Tabular Editor are always available for renaming, even
// when the chosen reports do not reference them.
foreach (var t in Selected.Tables) getOrAddReference("Table", t.Name, null);
foreach (var m in Selected.Measures) getOrAddReference("Measure", m.Table.Name, m.Name);
foreach (var c in Selected.Columns) getOrAddReference("Column", c.Table.Name, c.Name);

var references = referencesByKey.Values
    .OrderBy(r => r.Kind == "Table" ? 0 : r.Kind == "Measure" ? 1 : 2)
    .ThenBy(r => r.Table)
    .ThenBy(r => r.Field)
    .ToList();

if (references.Count == 0)
{
    Info("No table, measure, or column references were found in the selected reports. No changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 4: which references - the Tabular Editor selection, or pick from a tree
// ---------------------------------------------------------------------------
var selectedReferenceKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

if (useModelSelection)
{
    foreach (var key in modelSelectionKeys) selectedReferenceKeys.Add(key);
}
else
{
    using (var picker = new Form())
    {
        var layout = CreateDialogLayout(picker, "Select PBIR references to replace", 5);
        var font = new Font("Segoe UI", 9);
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblSearch = new Label {
            Text = "Search type, table, field, or model status (\"exists\" / \"missing\"). Tip: select objects in Tabular Editor before running to skip this step.",
            AutoSize = true, MaximumSize = new Size(1100, 0), Font = font
        };
        var txtSearch = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var tree = new TreeView { Dock = DockStyle.Fill, Font = new Font(font, FontStyle.Bold), CheckBoxes = true };
        var lblStatus = new Label { Text = "0 references selected", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var missingButton = new Button { Text = "Check missing", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var checkVisibleButton = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearButton = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { missingButton, checkVisibleButton, clearButton, okButton, cancelButton });

        Action updateStatus = () => lblStatus.Text = selectedReferenceKeys.Count + " reference(s) selected";
        bool suppress = false;
        Action<TreeNode, bool> setNodeAndDescendants = null;
        setNodeAndDescendants = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var reference = node.Tag as GeneralReference;
            if (reference != null)
            {
                if (isChecked) selectedReferenceKeys.Add(reference.Key);
                else selectedReferenceKeys.Remove(reference.Key);
            }
            foreach (TreeNode child in node.Nodes) setNodeAndDescendants(child, isChecked);
        };

        Func<GeneralReference, string> referenceText = reference =>
            (reference.Kind == "Table" ? reference.Table : reference.Field) +
            "  [" + (reference.ExistsInModel ? "exists in model" : "MISSING from model") +
            "; " + reference.Occurrences + " occurrence(s); " + reference.Files.Count + " file(s)]";

        Action rebuildTree = () =>
        {
            suppress = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();
            var search = (txtSearch.Text ?? "").Trim();
            var filtered = references.Where(r =>
                string.IsNullOrEmpty(search)
                || r.Kind.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || r.Table.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || (r.Field ?? "").IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || (r.ExistsInModel ? "exists" : "missing").IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();

            // Tables > Table   |   Measures / Columns > Table > Field
            foreach (var kindGroup in filtered.GroupBy(r => r.Kind))
            {
                var kindNode = new TreeNode(kindGroup.Key + "s");
                tree.Nodes.Add(kindNode);

                if (kindGroup.Key == "Table")
                {
                    foreach (var reference in kindGroup)
                        kindNode.Nodes.Add(new TreeNode(referenceText(reference)) {
                            Tag = reference, NodeFont = new Font("Segoe UI", 9),
                            Checked = selectedReferenceKeys.Contains(reference.Key)
                        });
                    continue;
                }

                foreach (var tableGroup in kindGroup.GroupBy(r => r.Table))
                {
                    var tableNode = new TreeNode(tableGroup.Key) { NodeFont = new Font("Segoe UI", 9) };
                    kindNode.Nodes.Add(tableNode);
                    foreach (var reference in tableGroup)
                        tableNode.Nodes.Add(new TreeNode(referenceText(reference)) {
                            Tag = reference, NodeFont = new Font("Segoe UI", 9),
                            Checked = selectedReferenceKeys.Contains(reference.Key)
                        });
                }
            }
            tree.ExpandAll();
            foreach (TreeNode kindNode in tree.Nodes)
            {
                bool allChecked = true;
                foreach (TreeNode child in kindNode.Nodes)
                {
                    if (child.Nodes.Count > 0) child.Checked = child.Nodes.Cast<TreeNode>().All(n => n.Checked);
                    if (!child.Checked) allChecked = false;
                }
                kindNode.Checked = allChecked && kindNode.Nodes.Count > 0;
            }
            if (tree.Nodes.Count > 0) tree.Nodes[0].EnsureVisible();
            tree.EndUpdate();
            suppress = false;
            updateStatus();
        };

        tree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;
            setNodeAndDescendants(e.Node, e.Node.Checked);
            suppress = false;
            updateStatus();
        };
        txtSearch.TextChanged += (s, e) => rebuildTree();
        checkVisibleButton.Click += (s, e) =>
        {
            suppress = true;
            foreach (TreeNode node in tree.Nodes) setNodeAndDescendants(node, true);
            suppress = false;
            updateStatus();
        };
        missingButton.Click += (s, e) =>
        {
            foreach (var reference in references.Where(r => !r.ExistsInModel))
                selectedReferenceKeys.Add(reference.Key);
            rebuildTree();
        };
        clearButton.Click += (s, e) =>
        {
            selectedReferenceKeys.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (selectedReferenceKeys.Count == 0)
            {
                MessageBox.Show("Select at least one reference.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            picker.DialogResult = DialogResult.OK;
            picker.Close();
        };

        picker.AcceptButton = okButton;
        picker.CancelButton = cancelButton;
        layout.Controls.Add(lblSearch, 0, 0);
        layout.Controls.Add(txtSearch, 0, 1);
        layout.Controls.Add(tree, 0, 2);
        layout.Controls.Add(lblStatus, 0, 3);
        layout.Controls.Add(buttons, 0, 4);
        rebuildTree();
        if (picker.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }
}

var selectedReferences = references.Where(r => selectedReferenceKeys.Contains(r.Key)).ToList();

// ---------------------------------------------------------------------------
// Step 5: one editable grid - existing name -> new name per reference
// ---------------------------------------------------------------------------

// Describes what a new name will do, for the live Result column.
Func<GeneralReference, string, string> describeAction = (reference, newName) =>
{
    newName = (newName ?? "").Trim();
    if (newName.Length == 0) return "Enter a new name";
    var oldName = reference.Kind == "Table" ? reference.Table : reference.Field;
    if (string.Equals(newName, oldName, StringComparison.Ordinal)) return "Unchanged";

    if (reference.Kind == "Table")
    {
        var src = Model.Tables.FirstOrDefault(t => SameName(t.Name, reference.Table));
        var tgt = Model.Tables.FirstOrDefault(t => SameName(t.Name, newName));
        if (tgt != null && tgt != src) return "PBIR only - repoint to existing table";
        if (src == null) return "PBIR only - source table not in model";
        return "Rename in model + PBIR";
    }

    if (reference.Kind == "Measure")
    {
        var src = Model.AllMeasures.FirstOrDefault(m => SameName(m.Table.Name, reference.Table) && SameName(m.Name, reference.Field));
        var tgt = Model.AllMeasures.FirstOrDefault(m => m != src && SameName(m.Name, newName));
        if (tgt != null) return "PBIR only - repoint to existing measure (" + tgt.Table.Name + ")";
        if (src == null) return "PBIR only - source measure not in model";
        return "Rename in model + PBIR";
    }

    var srcCol = Model.AllColumns.FirstOrDefault(c => SameName(c.Table.Name, reference.Table) && SameName(c.Name, reference.Field));
    var tgtCol = Model.AllColumns.FirstOrDefault(c => c != srcCol && SameName(c.Table.Name, reference.Table) && SameName(c.Name, newName));
    if (tgtCol != null) return "PBIR only - repoint to existing column";
    if (srcCol == null) return "PBIR only - source column not in model";
    return "Rename in model + PBIR";
};

var replacements = new List<GeneralReplacement>();

using (var form = new Form())
{
    var layout = CreateDialogLayout(form, "Set new names", 4);
    var font = new Font("Segoe UI", 9);
    var boldFont = new Font("Segoe UI", 9, FontStyle.Bold);
    var editBack = Color.FromArgb(255, 244, 179);   // highlighted editable cells
    var editHeaderBack = Color.FromArgb(255, 214, 102);
    var readOnlyFore = Color.FromArgb(90, 90, 90);
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var lblInfo = new Label {
        Text = "Set new name: edit the highlighted column. " + selectedReferences.Count + " reference(s)" +
            (useModelSelection ? " from the Tabular Editor selection." : " selected."),
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
    };
    var grid = new DataGridView {
        Dock = DockStyle.Fill,
        AutoGenerateColumns = false,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.CellSelect,
        Font = font,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
        BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false,   // required for custom header colours
        EditMode = DataGridViewEditMode.EditOnEnter
    };
    grid.ColumnHeadersDefaultCellStyle.Font = boldFont;
    grid.DefaultCellStyle.ForeColor = readOnlyFore;

    grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Include", HeaderText = "Include", Width = 60 });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Kind", HeaderText = "Type", Width = 70, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Table", HeaderText = "Table", Width = 170, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Existing", HeaderText = "Existing name", Width = 220, ReadOnly = true });
    var targetColumn = new DataGridViewTextBoxColumn {
        Name = "Target", HeaderText = "Set new name (edit)", Width = 240, ReadOnly = false,
        DefaultCellStyle = new DataGridViewCellStyle {
            BackColor = editBack, ForeColor = Color.Black, Font = boldFont,
            SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black
        }
    };
    targetColumn.HeaderCell.Style.BackColor = editHeaderBack;
    targetColumn.HeaderCell.Style.ForeColor = Color.Black;
    grid.Columns.Add(targetColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn {
        Name = "Result", HeaderText = "Result", ReadOnly = true,
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 220
    });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "InModel", HeaderText = "In model", Width = 70, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Uses", HeaderText = "PBIR uses", Width = 75, ReadOnly = true });

    foreach (var reference in selectedReferences)
    {
        var existingName = reference.Kind == "Table" ? reference.Table : reference.Field;
        var index = grid.Rows.Add(
            true,
            reference.Kind,
            reference.Kind == "Table" ? "" : reference.Table,
            existingName,
            existingName,
            "",
            reference.ExistsInModel ? "Yes" : "MISSING",
            reference.Occurrences);
        grid.Rows[index].Tag = reference;
        if (!reference.ExistsInModel) grid.Rows[index].Cells["InModel"].Style.ForeColor = Color.DarkRed;
    }

    Action<int> updateResult = rowIndex =>
    {
        var row = grid.Rows[rowIndex];
        var reference = (GeneralReference)row.Tag;
        var cell = row.Cells["Result"];
        if (!Convert.ToBoolean(row.Cells["Include"].Value))
        {
            cell.Value = "Skip - left unchanged";
            cell.Style.ForeColor = Color.Gray;
            return;
        }
        var text = describeAction(reference, Convert.ToString(row.Cells["Target"].Value));
        cell.Value = text;
        cell.Style.ForeColor =
            text.StartsWith("Rename") ? Color.DarkGreen :
            text.StartsWith("PBIR only") ? Color.DarkOrange :
            text == "Unchanged" ? Color.Gray : Color.DarkRed;
    };
    for (int i = 0; i < grid.Rows.Count; i++) updateResult(i);
    grid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) updateResult(e.RowIndex); };
    grid.CurrentCellDirtyStateChanged += (s, e) =>
    {
        if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell)
            grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
    };

    // Autocomplete existing model names of the same type while typing, so
    // repointing to an existing object is easy.
    grid.EditingControlShowing += (s, e) =>
    {
        var textBox = e.Control as TextBox;
        if (textBox == null) return;
        if (grid.CurrentCell == null || grid.CurrentCell.OwningColumn.Name != "Target")
        {
            textBox.AutoCompleteMode = AutoCompleteMode.None;
            return;
        }
        var reference = (GeneralReference)grid.CurrentRow.Tag;
        var names = reference.Kind == "Table"
            ? Model.Tables.Select(t => t.Name)
            : reference.Kind == "Measure"
                ? Model.AllMeasures.Select(m => m.Name)
                : Model.AllColumns.Where(c => SameName(c.Table.Name, reference.Table)).Select(c => c.Name);
        var source = new AutoCompleteStringCollection();
        source.AddRange(names.Distinct().ToArray());
        textBox.AutoCompleteCustomSource = source;
        textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
        textBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
    };

    form.Shown += (s, e) =>
    {
        if (grid.Rows.Count == 0) return;
        grid.CurrentCell = grid.Rows[0].Cells["Target"];
        grid.BeginEdit(true);
    };

    var lblHint = new Label {
        Text = "Only the highlighted column is editable; existing model names are suggested while typing. A new name that already exists in the model repoints PBIR to that object (no model rename). " +
            "Measures and columns stay on their table; renaming a table in the same run is taken into account. Uncheck Include to leave a row untouched.",
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
    };
    var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var okButton = new Button { Text = "Apply", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { okButton, cancelButton });

    okButton.Click += (s, e) =>
    {
        grid.EndEdit();
        replacements.Clear();
        foreach (DataGridViewRow row in grid.Rows)
        {
            if (!Convert.ToBoolean(row.Cells["Include"].Value)) continue;
            var reference = (GeneralReference)row.Tag;
            var newName = (Convert.ToString(row.Cells["Target"].Value) ?? "").Trim();
            if (newName.Length == 0)
            {
                MessageBox.Show("Enter a new name for " + reference.Kind.ToLower() + " " + reference.Label + ", or uncheck Include.",
                    "Blank name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (reference.Kind == "Table")
            {
                if (string.Equals(newName, reference.Table, StringComparison.Ordinal)) continue;
                replacements.Add(new GeneralReplacement { Kind = "Table", OldTable = reference.Table, NewTable = newName });
            }
            else
            {
                if (string.Equals(newName, reference.Field, StringComparison.Ordinal)) continue;
                // Measures are repointed to the table that hosts an existing
                // measure with the new name; otherwise the field stays on its table.
                var newTable = reference.Table;
                if (reference.Kind == "Measure")
                {
                    var src = Model.AllMeasures.FirstOrDefault(m => SameName(m.Table.Name, reference.Table) && SameName(m.Name, reference.Field));
                    var tgt = Model.AllMeasures.FirstOrDefault(m => m != src && SameName(m.Name, newName));
                    if (tgt != null) newTable = tgt.Table.Name;
                }
                replacements.Add(new GeneralReplacement {
                    Kind = reference.Kind,
                    OldTable = reference.Table,
                    OldField = reference.Field,
                    NewTable = newTable,
                    NewField = newName
                });
            }
        }

        if (replacements.Count == 0)
        {
            MessageBox.Show("No new names were entered. Edit at least one highlighted cell, or Cancel.", "Nothing to change", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        form.DialogResult = DialogResult.OK;
        form.Close();
    };

    // No AcceptButton: Enter commits the cell being edited instead of closing.
    form.CancelButton = cancelButton;
    layout.Controls.Add(lblInfo, 0, 0);
    layout.Controls.Add(grid, 0, 1);
    layout.Controls.Add(lblHint, 0, 2);
    layout.Controls.Add(buttons, 0, 3);

    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

var tableMap = replacements
    .Where(r => r.Kind == "Table")
    .ToDictionary(r => r.OldTable, r => r.NewTable, StringComparer.OrdinalIgnoreCase);

// Compose field mappings with a table mapping selected in the same run: PBIR
// must use the table's final mapped name.
foreach (var replacement in replacements.Where(r => r.Kind != "Table"))
{
    string mappedTable;
    if (tableMap.TryGetValue(replacement.OldTable, out mappedTable)
        && string.Equals(replacement.NewTable, replacement.OldTable, StringComparison.OrdinalIgnoreCase))
        replacement.NewTable = mappedTable;
}

var fieldMap = replacements
    .Where(r => r.Kind != "Table")
    .ToDictionary(r => r.Key, r => r, StringComparer.OrdinalIgnoreCase);

// ---------------------------------------------------------------------------
// Step 6: classify model actions before mutating PBIR
// ---------------------------------------------------------------------------
var modelActions = new List<GeneralModelAction>();
var tableActionByOldName = new Dictionary<string, GeneralModelAction>(StringComparer.OrdinalIgnoreCase);

foreach (var replacement in replacements.Where(r => r.Kind == "Table"))
{
    var sourceTable = Model.Tables.FirstOrDefault(t => SameName(t.Name, replacement.OldTable));
    var targetTable = Model.Tables.FirstOrDefault(t => SameName(t.Name, replacement.NewTable));

    GeneralModelAction action;
    if (sourceTable == null)
    {
        action = new GeneralModelAction {
            ReplacementKey = replacement.Key,
            Kind = "Table",
            Action = "PBIR only",
            Detail = "Source table is missing from the open model"
        };
    }
    else if (targetTable != null && targetTable != sourceTable)
    {
        action = new GeneralModelAction {
            ReplacementKey = replacement.Key,
            Kind = "Table",
            Action = "PBIR remap",
            Detail = "Target table already exists; source table will not be overwritten"
        };
    }
    else
    {
        action = new GeneralModelAction {
            ReplacementKey = replacement.Key,
            Kind = "Table",
            Action = ActionRename,
            Detail = sourceTable.Name + " -> " + replacement.NewTable,
            TableObject = sourceTable,
            NewName = replacement.NewTable
        };
    }
    modelActions.Add(action);
    tableActionByOldName[replacement.OldTable] = action;
}

foreach (var replacement in replacements.Where(r => r.Kind != "Table"))
{
    var sourceMeasure = replacement.Kind == "Measure"
        ? Model.AllMeasures.FirstOrDefault(m => SameName(m.Table.Name, replacement.OldTable) && SameName(m.Name, replacement.OldField))
        : null;
    var sourceColumn = replacement.Kind == "Column"
        ? Model.AllColumns.FirstOrDefault(c => SameName(c.Table.Name, replacement.OldTable) && SameName(c.Name, replacement.OldField))
        : null;

    var sourceExists = sourceMeasure != null || sourceColumn != null;
    var effectiveSourceTable = replacement.OldTable;
    GeneralModelAction relatedTableAction;
    if (tableActionByOldName.TryGetValue(replacement.OldTable, out relatedTableAction)
        && relatedTableAction.Action == ActionRename)
        effectiveSourceTable = relatedTableAction.NewName;

    var remainsOnSourceTable =
        SameName(replacement.NewTable, replacement.OldTable)
        || SameName(replacement.NewTable, effectiveSourceTable);

    GeneralModelAction action;
    if (!sourceExists)
    {
        action = new GeneralModelAction {
            ReplacementKey = replacement.Key,
            Kind = replacement.Kind,
            Action = "PBIR only",
            Detail = "Source " + replacement.Kind.ToLower() + " is missing from the open model"
        };
    }
    else if (!remainsOnSourceTable)
    {
        action = new GeneralModelAction {
            ReplacementKey = replacement.Key,
            Kind = replacement.Kind,
            Action = "PBIR remap",
            Detail = "Target exists on table " + replacement.NewTable + "; model object will not be moved or renamed"
        };
    }
    else
    {
        var collision = replacement.Kind == "Measure"
            ? Model.AllMeasures.Any(m => m != sourceMeasure && SameName(m.Name, replacement.NewField))
            : sourceColumn.Table.Columns.Any(c => c != sourceColumn && SameName(c.Name, replacement.NewField));

        if (collision)
        {
            action = new GeneralModelAction {
                ReplacementKey = replacement.Key,
                Kind = replacement.Kind,
                Action = "PBIR remap",
                Detail = "Target name already exists; source object will not be overwritten"
            };
        }
        else
        {
            action = new GeneralModelAction {
                ReplacementKey = replacement.Key,
                Kind = replacement.Kind,
                Action = ActionRename,
                Detail = replacement.OldField + " -> " + replacement.NewField,
                MeasureObject = sourceMeasure,
                ColumnObject = sourceColumn,
                NewName = replacement.NewField
            };
        }
    }
    modelActions.Add(action);
}

var duplicateTableTargets = modelActions
    .Where(a => a.Kind == "Table" && a.Action == ActionRename)
    .GroupBy(a => a.NewName, StringComparer.OrdinalIgnoreCase)
    .Where(g => g.Count() > 1)
    .Select(g => g.Key)
    .ToList();
if (duplicateTableTargets.Count > 0)
    Error("Multiple source tables would be renamed to the same target name: " + string.Join(", ", duplicateTableTargets) + "\n\nNo changes made.");

var duplicateFieldTargets = replacements
    .Where(r => r.Kind != "Table"
        && modelActions.Any(a => a.ReplacementKey == r.Key && a.Action == ActionRename))
    .GroupBy(r => r.Kind == "Measure" ? "Measure|" + r.NewField : r.Kind + "|" + r.NewTable + "|" + r.NewField, StringComparer.OrdinalIgnoreCase)
    .Where(g => g.Count() > 1)
    .Select(g => g.Key)
    .ToList();
if (duplicateFieldTargets.Count > 0)
    Error("Multiple source fields would be renamed to the same target: " + string.Join(", ", duplicateFieldTargets) + "\n\nNo changes made.");

// ---------------------------------------------------------------------------
// Step 7: mutate references in memory
// ---------------------------------------------------------------------------
Func<JToken, string, Dictionary<string, string>, GeneralReplacement, int> mutate = null;
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
                    && SameName(entity, forcedFilterReplacement.OldTable)
                    && !SameName(entity, forcedFilterReplacement.NewTable))
                {
                    from["Entity"] = forcedFilterReplacement.NewTable;
                    entity = forcedFilterReplacement.NewTable;
                    changes++;
                }
                localAliases[alias] = entity;
            }
        }

        GeneralReplacement localForced = forcedFilterReplacement;
        var outerMeasure = obj["field"]?["Measure"];
        var outerColumn = obj["field"]?["Column"];
        var outerField = outerMeasure ?? outerColumn;
        if (outerField != null && obj["filter"] != null)
        {
            var kind = outerMeasure != null ? "Measure" : "Column";
            var table = (string)outerField["Expression"]?["SourceRef"]?["Entity"];
            var field = (string)outerField["Property"];
            GeneralReplacement hit;
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

            GeneralReplacement fieldHit = null;
            if (table != null && field != null)
            {
                fieldMap.TryGetValue(parentKey + "|" + table + "|" + field, out fieldHit);
                if (fieldHit == null)
                {
                    fieldHit = fieldMap.Values.FirstOrDefault(r =>
                        r.Kind == parentKey
                        && SameName(r.OldField, field)
                        && tableMap.ContainsKey(r.OldTable)
                        && SameName(tableMap[r.OldTable], table));
                }
            }
            if (fieldHit == null && localForced != null
                && localForced.Kind == parentKey
                && (SameName(localForced.OldTable, table) || SameName(localForced.NewTable, table)))
                fieldHit = localForced;

            if (fieldHit != null)
            {
                if (entityToken != null && !SameName((string)entityToken, fieldHit.NewTable))
                {
                    sourceRef["Entity"] = fieldHit.NewTable;
                    changes++;
                }
                if (!string.Equals(field, fieldHit.NewField, StringComparison.Ordinal))
                {
                    obj["Property"] = fieldHit.NewField;
                    changes++;
                }
            }
            else if (table != null && entityToken != null)
            {
                string replacementTable;
                if (tableMap.TryGetValue(table, out replacementTable)
                    && !SameName(table, replacementTable))
                {
                    sourceRef["Entity"] = replacementTable;
                    changes++;
                }
            }
        }

        var queryRef = obj["queryRef"];
        if (queryRef != null && queryRef.Type == JTokenType.String)
        {
            var value = (string)queryRef;
            var fieldHit = fieldMap.Values.FirstOrDefault(r => SameName(value, r.OldTable + "." + r.OldField));
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
            var fieldHit = fieldMap.Values.FirstOrDefault(r => SameName(value, r.OldTable + "." + r.OldField));
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

int plannedModelRenames = modelActions.Count(a => a.Action == ActionRename);

if (filesToWrite.Count == 0 && plannedModelRenames == 0)
{
    Info("The chosen new names did not produce any model or PBIR changes. No changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 8: preview mappings, model actions, and affected files
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
preview.Columns.Add("Scope");
preview.Columns.Add("Type");
preview.Columns.Add("Old reference");
preview.Columns.Add("New reference");
preview.Columns.Add("Model action");
preview.Columns.Add("Detail");

foreach (var replacement in replacements)
{
    var source = referencesByKey[replacement.Key];
    var modelAction = modelActions.First(a => a.ReplacementKey == replacement.Key);
    preview.Rows.Add(
        "Mapping",
        replacement.Kind,
        replacement.Kind == "Table" ? replacement.OldTable : replacement.OldTable + "." + replacement.OldField,
        replacement.Kind == "Table" ? replacement.NewTable : replacement.NewTable + "." + replacement.NewField,
        modelAction.Action,
        modelAction.Detail + " | " + source.Occurrences + " PBIR occurrence(s)");
}

foreach (var file in filesToWrite)
{
    var relativePath = file.Item2.Substring(file.Item1.Length).TrimStart('\\', '/');
    preview.Rows.Add(
        file.Item1 + " | " + relativePath,
        "PBIR file",
        "",
        "",
        "",
        "Update " + file.Item4 + " reference value(s)");
}

if (!ShowGrid(
        "Preview TOM and PBIR replacements",
        "Review " + plannedModelRenames + " model rename(s) and " + filesToWrite.Count + " PBIR file update(s). Select Continue to confirm and apply the changes.",
        preview,
        "Continue"))
{
    Info("Cancelled - no changes made.");
    return;
}

var confirm = MessageBox.Show(
    "Apply " + replacements.Count + " mapping(s): " + plannedModelRenames + " model rename(s) and " +
    filesToWrite.Count + " PBIR JSON file(s)?\n\n" +
    "Display names are not changed. Mappings to existing objects only repoint PBIR and do not overwrite model objects.\n" +
    "PBIR edits cannot be undone in Tabular Editor. If the model/reports are under source control, commit/sync first if required.\n\n" +
    "After running:\n" +
    (plannedModelRenames > 0
        ? "1. First save the model changes in Tabular Editor, and deploy/update the model in the service if required (e.g. live-connected reports). Updating PBIR alone is not sufficient.\n" +
          "2. Then apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports."
        : "Apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports."),
    "Confirm TOM and PBIR replacements",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no changes made."); return; }

// ---------------------------------------------------------------------------
// Step 9: rename eligible model objects, then write PBIR files
// ---------------------------------------------------------------------------
int modelRenamesApplied = 0;

// Rename fields before tables so their source table lookup remains stable.
foreach (var action in modelActions.Where(a => a.Action == ActionRename && a.Kind != "Table"))
{
    if (action.MeasureObject != null) action.MeasureObject.Name = action.NewName;
    else if (action.ColumnObject != null) action.ColumnObject.Name = action.NewName;
    modelRenamesApplied++;
}
foreach (var action in modelActions.Where(a => a.Action == ActionRename && a.Kind == "Table"))
{
    action.TableObject.Name = action.NewName;
    modelRenamesApplied++;
}

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
            "The update stopped while writing:\n{0}\n\n" +
            "Changes already made before the error:\n" +
            "- {1} model object rename(s)\n" +
            "- {2} PBIR reference value(s)\n" +
            "- {3} PBIR JSON file(s)\n\n" +
            "Error: {4}",
            file.Item2,
            modelRenamesApplied,
            referencesWritten,
            filesWritten,
            ex.Message));
    }
}

Info(string.Format(
    "Applied {0} mapping(s), renamed {1} model object(s), and updated {2} PBIR reference value(s) across {3} file(s).\n\n" +
    "Next:\n{4}",
    replacements.Count,
    modelRenamesApplied,
    referencesWritten,
    filesWritten,
    modelRenamesApplied > 0
        ? "1. First save the model changes in Tabular Editor, and deploy/update the model in the service if required (e.g. live-connected reports).\n" +
          "2. Then apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports."
        : "Apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports."));
