// ============================================================================
// 12 - Set display names for selected fields in PBIR.cs
//
// Sets explicit PBIR projection displayName values for selected semantic-model
// measures and columns across one or more reports.
//
// What it does:
//   1. Uses measures/columns already selected in Tabular Editor. If none are
//      selected, shows a searchable hierarchy:
//        Table > Display folder > Measures/Columns > Object
//   2. Asks you to pick a folder, finds every definition.pbir below it, and
//      lets you select which reports to include.
//   3. Scans all visual.json files in every selected report for matching
//      projections and discovers existing displayName values.
//   4. Lets you apply to all matching visuals or select visuals from a tree:
//        Report > Page title > Visual type > Visual title (visual ID if no title)
//      Report nodes start collapsed; checking a report checks all its visuals.
//   5. Shows an editable grid with one row per selected field. If a field
//      already has display names in the reports, the most commonly used value
//      is proposed. Otherwise, the model object name is proposed.
//   6. Shows a complete preview before writing any PBIR files.
//
// Only projection.displayName is changed. nativeQueryRef is deliberately
// preserved because it is also a generated query alias and may contain
// uniqueness suffixes such as "Year1". queryRef and field references are not
// modified by this macro.
//
// IMPORTANT:
//   - PBIR edits are not undoable through Tabular Editor's undo stack. If the
//     reports are under source control, commit/sync first if required.
//   - After the update, apply external changes using Desktop Bridge, or close
//     and re-open the PBIR file in Power BI Desktop to see the changes in the
//     reports.
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

class DisplayField
{
    public string Kind;
    public string Table;
    public string Name;
    public string DisplayFolder;
    public string Key { get { return Kind + "|" + Table + "|" + Name; } }
    public string QualifiedName { get { return Table + "." + Name; } }
}

class DisplayPlan
{
    public DisplayField Field;
    public string ExistingValues;
    public string TargetDisplayName;
}

class DisplayVisual
{
    public string DefinitionFolder;
    public string File;
    public JObject Root;
    public string ReportName;
    public string PageDisplay;
    public string VisualType;
    public string VisualLabel;
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

// Standard resizable dialog: header rows (AutoSize), one fill row, and a
// bottom button row. Returns the layout panel; rows are added by the caller.
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
// Step 1: use selected fields, or show a searchable multi-select tree
// ---------------------------------------------------------------------------
var allFields = new List<DisplayField>();
allFields.AddRange(Model.AllMeasures.Select(m => new DisplayField {
    Kind = "Measure", Table = m.Table.Name, Name = m.Name, DisplayFolder = m.DisplayFolder ?? ""
}));
allFields.AddRange(Model.AllColumns.Select(c => new DisplayField {
    Kind = "Column", Table = c.Table.Name, Name = c.Name, DisplayFolder = c.DisplayFolder ?? ""
}));

if (allFields.Count == 0) Error("The model has no measures or columns.");

var selectedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
foreach (var m in Selected.Measures)
    selectedKeys.Add("Measure|" + m.Table.Name + "|" + m.Name);
foreach (var c in Selected.Columns)
    selectedKeys.Add("Column|" + c.Table.Name + "|" + c.Name);

var selectedFields = allFields.Where(f => selectedKeys.Contains(f.Key)).ToList();

if (selectedFields.Count == 0)
{
    var pickedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    using (var picker = new Form())
    {
        var layout = CreateDialogLayout(picker, "Select measures and columns", 5);
        var font = new Font("Segoe UI", 9);
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblSearch = new Label { Text = "Search table, display folder, measure, or column:", AutoSize = true, Font = font };
        var txtSearch = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var tree = new TreeView { Dock = DockStyle.Fill, Font = font, CheckBoxes = true };
        var lblStatus = new Label { Text = "0 fields selected", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisibleButton = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearButton = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { checkVisibleButton, clearButton, okButton, cancelButton });

        Action updateStatus = () => lblStatus.Text = pickedKeys.Count + " field(s) selected";

        bool suppressTreeEvents = false;
        Action<TreeNode, bool> setNodeAndDescendants = null;
        setNodeAndDescendants = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var field = node.Tag as DisplayField;
            if (field != null)
            {
                if (isChecked) pickedKeys.Add(field.Key);
                else pickedKeys.Remove(field.Key);
            }
            foreach (TreeNode child in node.Nodes)
                setNodeAndDescendants(child, isChecked);
        };

        Action rebuildTree = () =>
        {
            suppressTreeEvents = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();

            var search = (txtSearch.Text ?? "").Trim();
            var filtered = allFields
                .Where(f => string.IsNullOrEmpty(search)
                    || f.Table.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || f.DisplayFolder.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || f.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || f.Kind.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(f => f.Table)
                .ThenBy(f => f.DisplayFolder)
                .ThenBy(f => f.Kind)
                .ThenBy(f => f.Name)
                .ToList();

            var nodeLookup = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in filtered)
            {
                var pathParts = new List<string>() { field.Table };
                if (!string.IsNullOrWhiteSpace(field.DisplayFolder))
                    pathParts.AddRange(field.DisplayFolder.Split('\\').Where(p => !string.IsNullOrWhiteSpace(p)));
                pathParts.Add(field.Kind == "Measure" ? "Measures" : "Columns");

                TreeNodeCollection level = tree.Nodes;
                string path = "";
                foreach (var part in pathParts)
                {
                    path = string.IsNullOrEmpty(path) ? part : path + "\\" + part;
                    TreeNode folderNode;
                    if (!nodeLookup.TryGetValue(path, out folderNode))
                    {
                        folderNode = new TreeNode(part);
                        level.Add(folderNode);
                        nodeLookup[path] = folderNode;
                    }
                    level = folderNode.Nodes;
                }

                level.Add(new TreeNode(field.Name) {
                    Tag = field,
                    Checked = pickedKeys.Contains(field.Key)
                });
            }

            tree.ExpandAll();
            tree.EndUpdate();
            suppressTreeEvents = false;
            updateStatus();
        };

        tree.AfterCheck += (s, e) =>
        {
            if (suppressTreeEvents) return;
            suppressTreeEvents = true;
            setNodeAndDescendants(e.Node, e.Node.Checked);
            suppressTreeEvents = false;
            updateStatus();
        };
        txtSearch.TextChanged += (s, e) => rebuildTree();
        checkVisibleButton.Click += (s, e) =>
        {
            suppressTreeEvents = true;
            foreach (TreeNode node in tree.Nodes)
                setNodeAndDescendants(node, true);
            suppressTreeEvents = false;
            updateStatus();
        };
        clearButton.Click += (s, e) =>
        {
            pickedKeys.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (pickedKeys.Count == 0)
            {
                MessageBox.Show("Select at least one measure or column.", "No fields selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    selectedFields = allFields.Where(f => pickedKeys.Contains(f.Key))
        .OrderBy(f => f.Table).ThenBy(f => f.Kind).ThenBy(f => f.Name).ToList();
}

// ---------------------------------------------------------------------------
// Step 2: select root folder and PBIR reports
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
        Text = "Visuals in the selected reports that use the selected measures/columns can get new display names. Only visual.json projection display names are changed.",
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
// Step 3: discover visuals containing selected fields
// ---------------------------------------------------------------------------
var fieldByKey = selectedFields.ToDictionary(f => f.Key, f => f, StringComparer.OrdinalIgnoreCase);
var sourceFiles = new List<DisplayVisual>();

Func<string, string> literalText = raw =>
{
    if (string.IsNullOrEmpty(raw)) return null;
    return raw.Length >= 2 && raw.StartsWith("'") && raw.EndsWith("'")
        ? raw.Substring(1, raw.Length - 2).Replace("''", "'")
        : raw;
};

foreach (var reportFolder in reportFolders)
{
    var definitionFolder = Path.Combine(reportFolder, "definition");
    if (!Directory.Exists(definitionFolder))
    {
        Warning("Skipping selected report because its definition folder was not found: " + reportFolder);
        continue;
    }

    var pagesFolder = Path.Combine(definitionFolder, "pages");
    if (!Directory.Exists(pagesFolder))
    {
        Warning("Skipping selected report because definition\\pages was not found: " + reportFolder);
        continue;
    }

    foreach (var pageFolder in Directory.GetDirectories(pagesFolder))
    {
        var pageFile = Path.Combine(pageFolder, "page.json");
        if (!File.Exists(pageFile)) continue;

        JObject pageRoot = null;
        try { pageRoot = JObject.Parse(File.ReadAllText(pageFile)); } catch { }
        var pageName = (string)pageRoot?["displayName"] ?? Path.GetFileName(pageFolder);
        var reportName = CleanReportName(reportFolder);

        var visualsFolder = Path.Combine(pageFolder, "visuals");
        if (!Directory.Exists(visualsFolder)) continue;

        foreach (var file in Directory.GetFiles(visualsFolder, "visual.json", SearchOption.AllDirectories))
        {
            JObject root;
            try
            {
                root = JObject.Parse(File.ReadAllText(file));
            }
            catch (Exception ex)
            {
                Warning("Skipping unreadable visual.json:\n" + file + "\n\n" + ex.Message);
                continue;
            }

            var queryState = root["visual"]?["query"]?["queryState"] as JObject;
            if (queryState == null) continue;

            bool containsSelectedField = false;
            foreach (var role in queryState.Properties())
            {
                var projections = role.Value?["projections"] as JArray;
                if (projections == null) continue;

                foreach (var projection in projections.OfType<JObject>())
                {
                    var measure = projection["field"]?["Measure"];
                    var column = projection["field"]?["Column"];
                    var fieldToken = measure ?? column;
                    if (fieldToken == null) continue;

                    var kind = measure != null ? "Measure" : "Column";
                    var table = (string)fieldToken["Expression"]?["SourceRef"]?["Entity"];
                    var name = (string)fieldToken["Property"];
                    if (table != null && name != null && fieldByKey.ContainsKey(kind + "|" + table + "|" + name))
                    {
                        containsSelectedField = true;
                        break;
                    }
                }
                if (containsSelectedField) break;
            }

            if (!containsSelectedField) continue;

            var visualType = (string)root["visual"]?["visualType"] ?? "(unknown visual type)";
            var objectId = (string)root["name"] ?? Path.GetFileName(Path.GetDirectoryName(file));
            var title = literalText((string)root["visual"]?["visualContainerObjects"]?["title"]?[0]?["properties"]?["text"]?["expr"]?["Literal"]?["Value"]);
            var visualLabel = string.IsNullOrWhiteSpace(title) ? objectId : title;

            sourceFiles.Add(new DisplayVisual {
                DefinitionFolder = definitionFolder,
                File = file,
                Root = root,
                ReportName = reportName,
                PageDisplay = pageName,
                VisualType = visualType,
                VisualLabel = visualLabel
            });
        }
    }
}

if (sourceFiles.Count == 0)
{
    Info("None of the selected measures or columns are used in the selected reports.");
    return;
}

// ---------------------------------------------------------------------------
// Step 4: apply to all matching visuals, or select visuals from an event tree
// ---------------------------------------------------------------------------
bool selectSpecificVisuals = false;
using (var scopeForm = new Form())
{
    scopeForm.Text = "Choose visual scope";
    scopeForm.AutoSize = true;
    scopeForm.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    scopeForm.FormBorderStyle = FormBorderStyle.FixedDialog;
    scopeForm.MaximizeBox = false;
    scopeForm.MinimizeBox = false;
    scopeForm.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);

    var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
    var label = new Label {
        Text = sourceFiles.Count + " visual(s) contain at least one selected field.\nWhere should the display names be applied?",
        AutoSize = true, MaximumSize = new Size(520, 0), Font = font, Margin = new Padding(0, 0, 0, 12)
    };
    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    var allButton = new Button { Text = "Apply to all visuals", AutoSize = true, Font = font, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0) };
    var selectButton = new Button { Text = "Select visuals...", AutoSize = true, Font = font, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, Padding = new Padding(6, 3, 6, 3), DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { allButton, selectButton, cancelButton });
    layout.Controls.Add(label, 0, 0);
    layout.Controls.Add(buttonRow, 0, 1);

    allButton.Click += (s, e) =>
    {
        selectSpecificVisuals = false;
        scopeForm.DialogResult = DialogResult.OK;
        scopeForm.Close();
    };
    selectButton.Click += (s, e) =>
    {
        selectSpecificVisuals = true;
        scopeForm.DialogResult = DialogResult.OK;
        scopeForm.Close();
    };

    scopeForm.CancelButton = cancelButton;
    scopeForm.Controls.Add(layout);
    if (scopeForm.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

if (selectSpecificVisuals)
{
    var selectedVisualFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    using (var visualPicker = new Form())
    {
        var layout = CreateDialogLayout(visualPicker, "Select visuals", 5);
        var font = new Font("Segoe UI", 9);
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblSearch = new Label { Text = "Search report, page title, visual type, visual title, or visual ID (reports start collapsed; checking a report checks all its visuals):", AutoSize = true, Font = font };
        var txtSearch = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var tree = new TreeView { Dock = DockStyle.Fill, Font = new Font(font, FontStyle.Bold), CheckBoxes = true };
        var lblStatus = new Label { Text = "0 visuals selected", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisibleButton = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearButton = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { checkVisibleButton, clearButton, okButton, cancelButton });

        Action updateStatus = () => lblStatus.Text = selectedVisualFiles.Count + " visual(s) selected";
        bool suppress = false;

        Action<TreeNode, bool> setNodeAndDescendants = null;
        setNodeAndDescendants = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var visual = node.Tag as DisplayVisual;
            if (visual != null)
            {
                if (isChecked) selectedVisualFiles.Add(visual.File);
                else selectedVisualFiles.Remove(visual.File);
            }
            foreach (TreeNode child in node.Nodes)
                setNodeAndDescendants(child, isChecked);
        };

        // A parent node shows as checked when every visual below it is checked.
        // Needed because report nodes start collapsed.
        Func<TreeNode, bool> refreshParentChecks = null;
        refreshParentChecks = node =>
        {
            if (node.Nodes.Count == 0) return node.Checked;
            bool allChecked = true;
            foreach (TreeNode child in node.Nodes)
                if (!refreshParentChecks(child)) allChecked = false;
            node.Checked = allChecked;
            return allChecked;
        };

        // Report nodes the user has opened stay open across rebuilds
        // (search, Clear all). Everything below a report is always expanded.
        var expandedReports = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Action rebuildTree = () =>
        {
            suppress = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();

            var search = (txtSearch.Text ?? "").Trim();
            var filtered = sourceFiles.Where(v =>
                string.IsNullOrEmpty(search)
                || v.ReportName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.PageDisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualType.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualLabel.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || Path.GetFileName(Path.GetDirectoryName(v.File)).IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
            ).OrderBy(v => v.ReportName).ThenBy(v => v.PageDisplay).ThenBy(v => v.VisualType).ThenBy(v => v.VisualLabel).ToList();

            foreach (var reportGroup in filtered.GroupBy(v => v.ReportName))
            {
                // Report nodes inherit the tree's bold font; child nodes are set
                // back to regular. (A bold NodeFont on a regular-font tree gets
                // its text clipped, because node bounds use the tree's font.)
                var reportNode = new TreeNode(reportGroup.Key + "  (" + reportGroup.Count() + " visual(s))") { Name = reportGroup.Key };
                tree.Nodes.Add(reportNode);

                foreach (var pageGroup in reportGroup.GroupBy(v => v.PageDisplay))
                {
                    var pageNode = new TreeNode(pageGroup.Key) { NodeFont = font };
                    reportNode.Nodes.Add(pageNode);

                    foreach (var typeGroup in pageGroup.GroupBy(v => v.VisualType))
                    {
                        var typeNode = new TreeNode(typeGroup.Key) { NodeFont = font };
                        pageNode.Nodes.Add(typeNode);

                        foreach (var visual in typeGroup)
                        {
                            var visualNode = new TreeNode(visual.VisualLabel) {
                                Tag = visual,
                                NodeFont = font,
                                Checked = selectedVisualFiles.Contains(visual.File)
                            };
                            typeNode.Nodes.Add(visualNode);
                        }
                    }
                }

                refreshParentChecks(reportNode);
                reportNode.ExpandAll();
                if (string.IsNullOrEmpty(search) && !expandedReports.Contains(reportNode.Name))
                    reportNode.Collapse(true);
            }

            tree.EndUpdate();
            suppress = false;
            updateStatus();
        };

        tree.AfterExpand += (s, e) => { if (!suppress && e.Node.Level == 0) expandedReports.Add(e.Node.Name); };
        tree.AfterCollapse += (s, e) => { if (!suppress && e.Node.Level == 0) expandedReports.Remove(e.Node.Name); };

        tree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;
            setNodeAndDescendants(e.Node, e.Node.Checked);
            foreach (TreeNode reportNode in tree.Nodes) refreshParentChecks(reportNode);
            suppress = false;
            updateStatus();
        };
        txtSearch.TextChanged += (s, e) => rebuildTree();
        checkVisibleButton.Click += (s, e) =>
        {
            suppress = true;
            foreach (TreeNode node in tree.Nodes)
                setNodeAndDescendants(node, true);
            suppress = false;
            updateStatus();
        };
        clearButton.Click += (s, e) =>
        {
            selectedVisualFiles.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (selectedVisualFiles.Count == 0)
            {
                MessageBox.Show("Select at least one visual.", "No visuals selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            visualPicker.DialogResult = DialogResult.OK;
            visualPicker.Close();
        };

        visualPicker.AcceptButton = okButton;
        visualPicker.CancelButton = cancelButton;
        layout.Controls.Add(lblSearch, 0, 0);
        layout.Controls.Add(txtSearch, 0, 1);
        layout.Controls.Add(tree, 0, 2);
        layout.Controls.Add(lblStatus, 0, 3);
        layout.Controls.Add(buttons, 0, 4);
        rebuildTree();

        if (visualPicker.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }

    sourceFiles = sourceFiles.Where(v => selectedVisualFiles.Contains(v.File)).ToList();
}

// ---------------------------------------------------------------------------
// Step 5: collect existing display names within the selected visual scope
// ---------------------------------------------------------------------------
var existingNames = selectedFields.ToDictionary(
    f => f.Key,
    f => new List<string>(),
    StringComparer.OrdinalIgnoreCase);

foreach (var source in sourceFiles)
{
    var queryState = source.Root["visual"]?["query"]?["queryState"] as JObject;
    if (queryState == null) continue;
    foreach (var role in queryState.Properties())
    {
        var projections = role.Value?["projections"] as JArray;
        if (projections == null) continue;
        foreach (var projection in projections.OfType<JObject>())
        {
            var measure = projection["field"]?["Measure"];
            var column = projection["field"]?["Column"];
            var fieldToken = measure ?? column;
            if (fieldToken == null) continue;
            var kind = measure != null ? "Measure" : "Column";
            var table = (string)fieldToken["Expression"]?["SourceRef"]?["Entity"];
            var name = (string)fieldToken["Property"];
            var key = kind + "|" + table + "|" + name;
            if (!fieldByKey.ContainsKey(key)) continue;
            var displayName = (string)projection["displayName"];
            if (!string.IsNullOrWhiteSpace(displayName)) existingNames[key].Add(displayName);
        }
    }
}

var occurrenceCounts = selectedFields.ToDictionary(f => f.Key, f => 0, StringComparer.OrdinalIgnoreCase);
foreach (var source in sourceFiles)
{
    var queryState = source.Root["visual"]?["query"]?["queryState"] as JObject;
    if (queryState == null) continue;
    foreach (var role in queryState.Properties())
    {
        var projections = role.Value?["projections"] as JArray;
        if (projections == null) continue;
        foreach (var projection in projections.OfType<JObject>())
        {
            var measure = projection["field"]?["Measure"];
            var column = projection["field"]?["Column"];
            var fieldToken = measure ?? column;
            if (fieldToken == null) continue;
            var kind = measure != null ? "Measure" : "Column";
            var table = (string)fieldToken["Expression"]?["SourceRef"]?["Entity"];
            var name = (string)fieldToken["Property"];
            var key = kind + "|" + table + "|" + name;
            if (occurrenceCounts.ContainsKey(key)) occurrenceCounts[key]++;
        }
    }
}

var displayPlan = new List<DisplayPlan>();
foreach (var field in selectedFields)
{
    var distinctExisting = existingNames[field.Key]
        .GroupBy(n => n, StringComparer.OrdinalIgnoreCase)
        .OrderByDescending(g => g.Count())
        .ThenBy(g => g.Key)
        .ToList();

    displayPlan.Add(new DisplayPlan {
        Field = field,
        ExistingValues = distinctExisting.Count == 0
            ? "(none)"
            : string.Join(" | ", distinctExisting.Select(g => g.Key + " (" + g.Count() + ")")),
        TargetDisplayName = distinctExisting.Count == 0 ? field.Name : distinctExisting[0].Key
    });
}

// ---------------------------------------------------------------------------
// Step 6: edit proposed display names
// ---------------------------------------------------------------------------
using (var editor = new Form())
{
    var layout = CreateDialogLayout(editor, "Set display names", 4);
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
        Text = "Set display name: edit the highlighted column. The most common existing value is proposed; otherwise the model object name is used.",
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
    };
    var grid = new DataGridView {
        Dock = DockStyle.Fill,
        AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect,
        Font = font, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
        BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false,   // required for custom header colours
        EditMode = DataGridViewEditMode.EditOnEnter
    };
    grid.ColumnHeadersDefaultCellStyle.Font = boldFont;
    grid.DefaultCellStyle.ForeColor = readOnlyFore;

    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Kind", HeaderText = "Type", Width = 80, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Table", HeaderText = "Table", Width = 160, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Field", HeaderText = "Field", Width = 210, ReadOnly = true });
    var targetColumn = new DataGridViewTextBoxColumn {
        Name = "Target", HeaderText = "Set display name (edit)", Width = 260, ReadOnly = false,
        DefaultCellStyle = new DataGridViewCellStyle {
            BackColor = editBack, ForeColor = Color.Black, Font = boldFont,
            SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black
        }
    };
    targetColumn.HeaderCell.Style.BackColor = editHeaderBack;
    targetColumn.HeaderCell.Style.ForeColor = Color.Black;
    grid.Columns.Add(targetColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Occurrences", HeaderText = "Uses", Width = 55, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn {
        Name = "Existing", HeaderText = "Existing display names", ReadOnly = true,
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200,
        DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
    });

    foreach (var item in displayPlan)
        grid.Rows.Add(
            item.Field.Kind,
            item.Field.Table,
            item.Field.Name,
            item.TargetDisplayName,
            occurrenceCounts[item.Field.Key],
            item.ExistingValues);

    // Start with the cursor in the first editable cell
    editor.Shown += (s, e) =>
    {
        if (grid.Rows.Count == 0) return;
        grid.CurrentCell = grid.Rows[0].Cells["Target"];
        grid.BeginEdit(true);
    };

    var lblHint = new Label {
        Text = "Only the highlighted column is editable. Blank values are not allowed. Fields with 0 uses remain in the preview but do not modify files.",
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
    };
    var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var okButton = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { okButton, cancelButton });

    okButton.Click += (s, e) =>
    {
        grid.EndEdit();
        for (int i = 0; i < grid.Rows.Count; i++)
        {
            var value = Convert.ToString(grid.Rows[i].Cells["Target"].Value).Trim();
            if (string.IsNullOrEmpty(value))
            {
                MessageBox.Show("Display name cannot be blank for " + displayPlan[i].Field.QualifiedName + ".",
                    "Blank display name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            displayPlan[i].TargetDisplayName = value;
        }
        editor.DialogResult = DialogResult.OK;
        editor.Close();
    };

    editor.CancelButton = cancelButton;
    layout.Controls.Add(lblInfo, 0, 0);
    layout.Controls.Add(grid, 0, 1);
    layout.Controls.Add(lblHint, 0, 2);
    layout.Controls.Add(buttons, 0, 3);

    if (editor.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 7: build mutations and show preview before writing
// ---------------------------------------------------------------------------
var targetByKey = displayPlan.ToDictionary(
    p => p.Field.Key,
    p => p.TargetDisplayName,
    StringComparer.OrdinalIgnoreCase);

var filesToWrite = new List<Tuple<string, string, JObject, int>>();
var previewRows = new List<Tuple<string, string, string, string>>();

foreach (var source in sourceFiles)
{
    int changes = 0;
    var root = source.Root;
    var queryState = root["visual"]?["query"]?["queryState"] as JObject;
    if (queryState == null) continue;

    foreach (var role in queryState.Properties())
    {
        var projections = role.Value?["projections"] as JArray;
        if (projections == null) continue;

        foreach (var projection in projections.OfType<JObject>())
        {
            var measure = projection["field"]?["Measure"];
            var column = projection["field"]?["Column"];
            var fieldToken = measure ?? column;
            if (fieldToken == null) continue;

            var kind = measure != null ? "Measure" : "Column";
            var table = (string)fieldToken["Expression"]?["SourceRef"]?["Entity"];
            var name = (string)fieldToken["Property"];
            if (table == null || name == null) continue;

            var key = kind + "|" + table + "|" + name;
            string target;
            if (!targetByKey.TryGetValue(key, out target)) continue;

            var oldDisplayName = (string)projection["displayName"];
            if (string.Equals(oldDisplayName, target, StringComparison.Ordinal)) continue;

            projection["displayName"] = target;
            changes++;
            previewRows.Add(Tuple.Create(
                source.File,
                kind + " " + table + "." + name,
                string.IsNullOrEmpty(oldDisplayName) ? "(not set)" : oldDisplayName,
                target));
        }
    }

    if (changes > 0)
        filesToWrite.Add(Tuple.Create(source.DefinitionFolder, source.File, root, changes));
}

var preview = new System.Data.DataTable();
preview.Columns.Add("Visual file");
preview.Columns.Add("Field");
preview.Columns.Add("Existing display name");
preview.Columns.Add("New display name");

foreach (var row in previewRows)
{
    var source = sourceFiles.First(s => s.File == row.Item1);
    var relativePath = row.Item1.Substring(source.DefinitionFolder.Length).TrimStart('\\', '/');
    preview.Rows.Add(source.DefinitionFolder + " | " + relativePath, row.Item2, row.Item3, row.Item4);
}

foreach (var item in displayPlan.Where(p => occurrenceCounts[p.Field.Key] == 0))
    preview.Rows.Add("(not found in selected reports)", item.Field.Kind + " " + item.Field.QualifiedName, "(n/a)", item.TargetDisplayName);

if (filesToWrite.Count == 0)
{
    Output(preview);
    Info("No PBIR changes are required. The selected fields were not used, or all matching projections already have the requested display names.");
    return;
}

if (!ShowGrid(
        "Preview display-name changes",
        "Review the display-name changes per visual. Select Continue to confirm and write the changes.",
        preview,
        "Continue"))
{
    Info("Cancelled - no files were changed.");
    return;
}

var confirm = MessageBox.Show(
    "Update " + previewRows.Count + " projection display name(s) across " +
    filesToWrite.Count + " visual.json file(s) in " + reportFolders.Count + " report(s)?\n\n" +
    "PBIR edits cannot be undone in Tabular Editor. If the reports are under source control, commit/sync first if required.\n\n" +
    "Afterwards, apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    "Confirm display-name changes",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no files were changed."); return; }

// ---------------------------------------------------------------------------
// Step 8: write changed files and summarize any partial progress on failure
// ---------------------------------------------------------------------------
int filesWritten = 0;
int displayNamesWritten = 0;

foreach (var file in filesToWrite)
{
    try
    {
        File.WriteAllText(file.Item2, file.Item3.ToString(Newtonsoft.Json.Formatting.Indented));
        filesWritten++;
        displayNamesWritten += file.Item4;
    }
    catch (Exception ex)
    {
        Error(string.Format(
            "The update stopped while writing:\n{0}\n\n" +
            "Changes already made before the error:\n" +
            "- {1} projection display name(s)\n" +
            "- {2} visual.json file(s)\n\n" +
            "Error: {3}",
            file.Item2,
            displayNamesWritten,
            filesWritten,
            ex.Message));
    }
}

Info(string.Format(
    "Updated {0} projection display name(s) across {1} visual.json file(s) in {2} report(s).\n\n" +
    "Next: apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    displayNamesWritten,
    filesWritten,
    reportFolders.Count));
