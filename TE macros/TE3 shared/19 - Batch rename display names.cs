// ============================================================================
// 13 - Batch rename existing display names in PBIR.cs
//
// PBIR-only macro that standardizes complete projection.displayName values
// by their underlying semantic-model measure or column reference.
//
// What it does:
//   1. Asks you to pick a folder, finds every definition.pbir below it, and
//      lets you select which reports to include.
//   2. Discovers visuals containing explicit display names.
//   3. Lets you apply to all matching visuals or select from:
//        Report > Page title > Visual type > Visual title (visual ID if no title)
//      Report nodes start collapsed; checking a report checks all its visuals.
//   4. Groups every occurrence by the underlying measure/column reference.
//   5. Shows the original model field name, number of display-name variations,
//      examples with usage counts, and an editable display name to use.
//   6. Replaces the entire displayName value for every chosen occurrence.
//      If the display name to use is blank or equal to the original model
//      field name, the displayName property is REMOVED instead, so the visual
//      falls back to the field name.
//   7. Previews every change before writing.
//
// This macro changes projection.displayName only. It does not change the
// semantic model, field references, queryRef, or nativeQueryRef.
//
// IMPORTANT:
//   - Only PBIR files are changed. PBIR edits are not undoable through
//     Tabular Editor's undo stack. If the reports are under source control,
//     commit/sync first if required.
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

class RenameDisplayOccurrence
{
    public string Kind;
    public string Table;
    public string Field;
    public string OldDisplayName;
    public JObject Projection;
}

class RenameDisplayVisual
{
    public string DefinitionFolder;
    public string File;
    public JObject Root;
    public string ReportName;
    public string PageDisplay;
    public string VisualType;
    public string VisualLabel;
    public List<RenameDisplayOccurrence> Occurrences = new List<RenameDisplayOccurrence>();
}

class RenameDisplayPlan
{
    public string Key;
    public string Kind;
    public string Table;
    public string Field;
    public int OccurrenceCount;
    public int VariationCount;
    public string ExistingVariations;
    public string TargetDisplayName;
    public bool Include;
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

// A display name that is blank or equal to the model field name is removed,
// so the visual falls back to the field name.
Func<string, string, bool> IsRemoval = (target, field) =>
    string.IsNullOrWhiteSpace(target) || string.Equals(target.Trim(), field, StringComparison.Ordinal);

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
        Text = "Existing explicit display names in the selected reports can be standardized per measure/column. Only visual.json projection display names are changed; the semantic model is not affected.",
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
// Step 2: discover visuals and existing explicit displayName occurrences
// ---------------------------------------------------------------------------
var visuals = new List<RenameDisplayVisual>();

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
    var pagesFolder = Path.Combine(definitionFolder, "pages");
    if (!Directory.Exists(pagesFolder))
    {
        Warning("Skipping selected report because definition\\pages was not found: " + reportFolder);
        continue;
    }

    var reportName = CleanReportName(reportFolder);

    foreach (var pageFolder in Directory.GetDirectories(pagesFolder))
    {
        var pageFile = Path.Combine(pageFolder, "page.json");
        if (!File.Exists(pageFile)) continue;

        JObject pageRoot = null;
        try { pageRoot = JObject.Parse(File.ReadAllText(pageFile)); } catch { }
        var pageName = (string)pageRoot?["displayName"] ?? Path.GetFileName(pageFolder);

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

            var visual = new RenameDisplayVisual {
                DefinitionFolder = definitionFolder,
                File = file,
                Root = root,
                ReportName = reportName,
                PageDisplay = pageName,
                VisualType = (string)root["visual"]?["visualType"] ?? "(unknown visual type)"
            };

            var objectId = (string)root["name"] ?? Path.GetFileName(Path.GetDirectoryName(file));
            var title = literalText((string)root["visual"]?["visualContainerObjects"]?["title"]?[0]?["properties"]?["text"]?["expr"]?["Literal"]?["Value"]);
            visual.VisualLabel = string.IsNullOrWhiteSpace(title) ? objectId : title;

            var queryState = root["visual"]?["query"]?["queryState"] as JObject;
            if (queryState == null) continue;

            foreach (var role in queryState.Properties())
            {
                var projections = role.Value?["projections"] as JArray;
                if (projections == null) continue;

                foreach (var projection in projections.OfType<JObject>())
                {
                    var displayName = (string)projection["displayName"];
                    if (string.IsNullOrWhiteSpace(displayName)) continue;

                    var measure = projection["field"]?["Measure"];
                    var column = projection["field"]?["Column"];
                    var fieldToken = measure ?? column;
                    if (fieldToken == null) continue;

                    var table = (string)fieldToken["Expression"]?["SourceRef"]?["Entity"];
                    var field = (string)fieldToken["Property"];
                    if (table == null || field == null) continue;

                    visual.Occurrences.Add(new RenameDisplayOccurrence {
                        Kind = measure != null ? "Measure" : "Column",
                        Table = table,
                        Field = field,
                        OldDisplayName = displayName,
                        Projection = projection
                    });
                }
            }

            if (visual.Occurrences.Count > 0) visuals.Add(visual);
        }
    }
}

if (visuals.Count == 0)
{
    Info("No explicit projection displayName values were found in the selected reports. No changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 3: choose all visuals or a specific visual scope
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
        Text = visuals.Count + " visual(s) contain explicit display names.\nWhere should the batch rename be applied?",
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
    var selectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    using (var picker = new Form())
    {
        var layout = CreateDialogLayout(picker, "Select visuals containing display names", 5);
        var font = new Font("Segoe UI", 9);
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblSearch = new Label {
            Text = "Search report, page title, visual type, visual title, visual ID, field, or display name (reports start collapsed; checking a report checks all its visuals):",
            AutoSize = true, MaximumSize = new Size(1100, 0), Font = font
        };
        var txtSearch = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        // The tree font is bold so bold report names are not clipped; child
        // nodes are set back to the regular font.
        var tree = new TreeView { Dock = DockStyle.Fill, Font = new Font(font, FontStyle.Bold), CheckBoxes = true };
        var lblStatus = new Label { Text = "0 visuals selected", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisibleButton = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearButton = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { checkVisibleButton, clearButton, okButton, cancelButton });

        Action updateStatus = () => lblStatus.Text = selectedFiles.Count + " visual(s) selected";
        bool suppress = false;

        Action<TreeNode, bool> setNodeAndDescendants = null;
        setNodeAndDescendants = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var visual = node.Tag as RenameDisplayVisual;
            if (visual != null)
            {
                if (isChecked) selectedFiles.Add(visual.File);
                else selectedFiles.Remove(visual.File);
            }
            foreach (TreeNode child in node.Nodes)
                setNodeAndDescendants(child, isChecked);
        };

        // A parent node shows as checked when every visual below it is checked.
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

        // Report nodes the user has opened stay open across rebuilds.
        var expandedReports = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Action rebuildTree = () =>
        {
            suppress = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();

            var search = (txtSearch.Text ?? "").Trim();
            var filtered = visuals.Where(v =>
                string.IsNullOrEmpty(search)
                || v.ReportName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.PageDisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualType.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualLabel.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || Path.GetFileName(Path.GetDirectoryName(v.File)).IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.Occurrences.Any(o =>
                    o.Table.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || o.Field.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || o.OldDisplayName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
            ).OrderBy(v => v.ReportName).ThenBy(v => v.PageDisplay).ThenBy(v => v.VisualType).ThenBy(v => v.VisualLabel).ToList();

            foreach (var reportGroup in filtered.GroupBy(v => v.ReportName))
            {
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
                            var details = string.Join(", ", visual.Occurrences
                                .Select(o => o.OldDisplayName)
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .Take(3));
                            typeNode.Nodes.Add(new TreeNode(visual.VisualLabel + "  ->  " + details) {
                                Tag = visual,
                                NodeFont = font,
                                Checked = selectedFiles.Contains(visual.File)
                            });
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
            selectedFiles.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("Select at least one visual.", "No visuals selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    visuals = visuals.Where(v => selectedFiles.Contains(v.File)).ToList();
}

// ---------------------------------------------------------------------------
// Step 4: group occurrences by the underlying measure/column and choose the
// complete display name that should be used for each field
// ---------------------------------------------------------------------------
var allOccurrences = visuals.SelectMany(v => v.Occurrences).ToList();
var standardizationPlan = new List<RenameDisplayPlan>();

foreach (var group in allOccurrences.GroupBy(
    o => o.Kind + "|" + o.Table + "|" + o.Field,
    StringComparer.OrdinalIgnoreCase))
{
    var first = group.First();
    var variations = group
        .GroupBy(o => o.OldDisplayName, StringComparer.OrdinalIgnoreCase)
        .OrderByDescending(g => g.Count())
        .ThenByDescending(g => string.Equals(g.Key, first.Field, StringComparison.OrdinalIgnoreCase))
        .ThenBy(g => g.Key)
        .ToList();

    // Proposal logic:
    //   - one variation: keep that existing display name
    //   - multiple variations: use the most frequent one
    //   - tied frequency: prefer the original model field name when present
    var highestCount = variations.Max(g => g.Count());
    var highest = variations.Where(g => g.Count() == highestCount).ToList();
    var exactModelName = highest.FirstOrDefault(g =>
        string.Equals(g.Key, first.Field, StringComparison.OrdinalIgnoreCase));
    var proposed = exactModelName != null ? exactModelName.Key : highest.OrderBy(g => g.Key).First().Key;

    standardizationPlan.Add(new RenameDisplayPlan {
        Key = group.Key,
        Kind = first.Kind,
        Table = first.Table,
        Field = first.Field,
        OccurrenceCount = group.Count(),
        VariationCount = variations.Count,
        ExistingVariations = string.Join(" | ", variations.Select(g => g.Key + " (" + g.Count() + ")")),
        TargetDisplayName = proposed,
        Include = true
    });
}

standardizationPlan = standardizationPlan
    .OrderBy(p => p.Table)
    .ThenBy(p => p.Kind)
    .ThenBy(p => p.Field)
    .ToList();

using (var form = new Form())
{
    var layout = CreateDialogLayout(form, "Set display names", 4);
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
        Text = "Set display name: edit the highlighted column. " + allOccurrences.Count + " explicit display-name occurrence(s) grouped into " +
            standardizationPlan.Count + " measure/column reference(s).",
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
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Kind", HeaderText = "Type", Width = 75, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Table", HeaderText = "Table", Width = 150, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "OriginalName", HeaderText = "Original model field name", Width = 210, ReadOnly = true });
    var targetColumn = new DataGridViewTextBoxColumn {
        Name = "Target", HeaderText = "Set display name (edit)", Width = 240, ReadOnly = false,
        DefaultCellStyle = new DataGridViewCellStyle {
            BackColor = editBack, ForeColor = Color.Black, Font = boldFont,
            SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black
        }
    };
    targetColumn.HeaderCell.Style.BackColor = editHeaderBack;
    targetColumn.HeaderCell.Style.ForeColor = Color.Black;
    grid.Columns.Add(targetColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Result", HeaderText = "Result", Width = 190, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Uses", HeaderText = "Uses", Width = 50, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Variations", HeaderText = "Variations", Width = 70, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn {
        Name = "Examples", HeaderText = "Existing display names (usage count)", ReadOnly = true,
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200,
        DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
    });

    foreach (var item in standardizationPlan)
        grid.Rows.Add(
            item.Include,
            item.Kind,
            item.Table,
            item.Field,
            item.TargetDisplayName,
            "",
            item.OccurrenceCount,
            item.VariationCount,
            item.ExistingVariations);

    // Result column explains what will happen to each row
    Action<int> updateResult = rowIndex =>
    {
        var row = grid.Rows[rowIndex];
        bool include = Convert.ToBoolean(row.Cells["Include"].Value);
        var target = Convert.ToString(row.Cells["Target"].Value);
        var cell = row.Cells["Result"];
        if (!include)
        {
            cell.Value = "Skip - left unchanged";
            cell.Style.ForeColor = Color.Gray;
        }
        else if (IsRemoval(target, standardizationPlan[rowIndex].Field))
        {
            cell.Value = "Remove display name (use field name)";
            cell.Style.ForeColor = Color.DarkRed;
        }
        else
        {
            cell.Value = "Set display name";
            cell.Style.ForeColor = Color.DarkGreen;
        }
    };
    for (int i = 0; i < grid.Rows.Count; i++) updateResult(i);
    grid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) updateResult(e.RowIndex); };
    // Commit checkbox clicks immediately so the Result column updates
    grid.CurrentCellDirtyStateChanged += (s, e) =>
    {
        if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell)
            grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
    };

    form.Shown += (s, e) =>
    {
        if (grid.Rows.Count == 0) return;
        grid.CurrentCell = grid.Rows[0].Cells["Target"];
        grid.BeginEdit(true);
    };

    var lblHint = new Label {
        Text = "Only the highlighted column is editable. Leave it blank or equal to the original model field name to REMOVE the display name. Uncheck Include to leave a field untouched.",
        AutoSize = true, MaximumSize = new Size(1100, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
    };
    var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var okButton = new Button { Text = "Apply", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { okButton, cancelButton });

    okButton.Click += (s, e) =>
    {
        grid.EndEdit();
        int enabledRows = 0;
        for (int i = 0; i < grid.Rows.Count; i++)
        {
            bool include = Convert.ToBoolean(grid.Rows[i].Cells["Include"].Value);
            standardizationPlan[i].Include = include;
            if (!include) continue;

            enabledRows++;
            standardizationPlan[i].TargetDisplayName = (Convert.ToString(grid.Rows[i].Cells["Target"].Value) ?? "").Trim();
        }

        if (enabledRows == 0)
        {
            MessageBox.Show("Include at least one row.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

// ---------------------------------------------------------------------------
// Step 5: replace (or remove) displayName values and preview before writing
// ---------------------------------------------------------------------------
const string RemovedLabel = "(removed - uses field name)";
var filesToWrite = new List<Tuple<string, string, JObject, int>>();
var previewRows = new List<Tuple<string, string, string, string, string>>();
var planByField = standardizationPlan
    .Where(p => p.Include)
    .ToDictionary(p => p.Key, p => p, StringComparer.OrdinalIgnoreCase);

foreach (var visual in visuals)
{
    int changes = 0;
    foreach (var occurrence in visual.Occurrences)
    {
        RenameDisplayPlan plan;
        var key = occurrence.Kind + "|" + occurrence.Table + "|" + occurrence.Field;
        if (!planByField.TryGetValue(key, out plan)) continue;

        string newValue;
        if (IsRemoval(plan.TargetDisplayName, occurrence.Field))
        {
            occurrence.Projection.Remove("displayName");
            newValue = RemovedLabel;
        }
        else
        {
            if (string.Equals(plan.TargetDisplayName, occurrence.OldDisplayName, StringComparison.Ordinal)) continue;
            occurrence.Projection["displayName"] = plan.TargetDisplayName;
            newValue = plan.TargetDisplayName;
        }

        changes++;
        previewRows.Add(Tuple.Create(
            visual.File,
            occurrence.Kind,
            occurrence.Table + "." + occurrence.Field,
            occurrence.OldDisplayName,
            newValue));
    }

    if (changes > 0)
        filesToWrite.Add(Tuple.Create(visual.DefinitionFolder, visual.File, visual.Root, changes));
}

var preview = new System.Data.DataTable();
preview.Columns.Add("Visual file");
preview.Columns.Add("Type");
preview.Columns.Add("Field");
preview.Columns.Add("Existing display name");
preview.Columns.Add("New display name");

foreach (var row in previewRows)
{
    var visual = visuals.First(v => v.File == row.Item1);
    var relativePath = row.Item1.Substring(visual.DefinitionFolder.Length).TrimStart('\\', '/');
    preview.Rows.Add(
        visual.DefinitionFolder + " | " + relativePath,
        row.Item2,
        row.Item3,
        row.Item4,
        row.Item5);
}

if (filesToWrite.Count == 0)
{
    Info("No display-name changes are required. Every selected occurrence already uses its chosen display name. No changes made.");
    return;
}

int removals = previewRows.Count(r => r.Item5 == RemovedLabel);

if (!ShowGrid(
        "Preview display-name changes",
        "Review the display-name changes per visual" +
            (removals > 0 ? " (" + removals + " display name(s) will be removed so the field name is used)" : "") +
            ". Select Continue to confirm and write the changes.",
        preview,
        "Continue"))
{
    Info("Cancelled - no files were changed.");
    return;
}

var confirm = MessageBox.Show(
    "Update " + previewRows.Count + " display-name occurrence(s) across " +
    filesToWrite.Count + " visual.json file(s) in " + reportFolders.Count + " report(s)?" +
    (removals > 0 ? "\n(" + removals + " of these remove the display name so the field name is used.)" : "") + "\n\n" +
    "Only PBIR files are changed; the semantic model is not affected. PBIR edits cannot be undone in Tabular Editor. " +
    "If the reports are under source control, commit/sync first if required.\n\n" +
    "Afterwards, apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    "Confirm display-name changes",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no files were changed."); return; }

// ---------------------------------------------------------------------------
// Step 6: write files and summarize partial progress if an error occurs
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
            "- {1} display-name change(s)\n" +
            "- {2} visual.json file(s)\n\n" +
            "Error: {3}",
            file.Item2,
            displayNamesWritten,
            filesWritten,
            ex.Message));
    }
}

Info(string.Format(
    "Updated {0} display name(s) across {1} visual.json file(s) in {2} report(s){3}.\n\n" +
    "Only PBIR files were changed; the semantic model is not affected.\n" +
    "Next: apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    displayNamesWritten,
    filesWritten,
    reportFolders.Count,
    removals > 0 ? " (" + removals + " removed so the field name is used)" : ""));
