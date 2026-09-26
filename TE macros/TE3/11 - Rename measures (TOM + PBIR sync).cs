// ============================================================================
// 11 - Rename measures (TOM + PBIR sync).cs
//
// Combined macro: renames one or more measures in the semantic model (TOM)
// AND updates every reference to those measures in a PBIR report's files
// (query projections, filterConfig filters incl. nested advanced-filter
// Where/From clauses, and selector-scoped formatting entries) in one run.
//
// New names are derived with a single Find/Replace text pattern applied to
// every selected measure's current name (e.g. Find "Errors", Replace "Issues"
// turns "Testing Errors YTD" into "Testing Issues YTD").
//
// What it does:
//   1. Uses the measures already selected in Tabular Editor. If no measures
//      are selected, shows a searchable Table > Display folder > Measure tree.
//      You then enter the Find/Replace text.
//   2. Computes the new name for each selected measure. Measures whose name
//      doesn't contain the Find text are skipped (no change). Measures whose
//      computed new name already exists elsewhere in the model - or collides
//      with another selected measure's new name - are flagged and SKIPPED
//      (no automatic overwrite of an existing name).
//   3. Asks you to pick a folder, finds every definition.pbir below it, and
//      lets you select which reports to include. Each selected report is
//      scanned in full (all pages/visuals + report.json). The model is renamed
//      once; every chosen PBIR report is updated from the same old -> new plan.
//   4. Shows one combined preview grid (model renames + PBIR files to be
//      touched) and asks for confirmation before changing anything.
//   5. On confirm: renames the measures in the model first (Tabular Editor's
//      TOM engine automatically updates any dependent DAX expressions that
//      reference the renamed measure), then writes the updated PBIR files.
//
// IMPORTANT - read before running:
//   - Model renames are undoable with Ctrl+Z in Tabular Editor, but the PBIR
//     file edits made by this script are NOT. If the model/reports are under
//     source control, commit/sync first if required.
//   - Afterwards, FIRST save/deploy/sync the semantic-model changes. Updating
//     only the report files is not sufficient: the renamed measures must exist
//     in the model used by the report (for live connections, the model in the
//     service).
//   - Then apply external changes using Desktop Bridge, or close and re-open
//     the PBIR file in Power BI Desktop to see the changes in the reports.
//   - If OTHER reports live-connect to this same published semantic model,
//     this script does not touch their PBIR files - rerun it against each
//     affected report (or expect Desktop to show "field not found" until you
//     do, or until the fields are manually re-mapped in that report).
//
// Safe to re-run: measures that no longer match the Find text, or whose
// computed name already exists, are simply skipped on a later run.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

class PlanRow
{
    public Measure Measure;
    public string Table, OldName, NewName, Action;
}

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
// Step 1: use selected measures, or pick measures from a searchable tree,
// then collect the Find/Replace pattern
// ---------------------------------------------------------------------------
var allMeasures = Model.AllMeasures.OrderBy(m => m.Table.Name).ThenBy(m => m.Name).ToList();
if (allMeasures.Count == 0) Error("The model has no measures.");

var selectedMeasures = Selected.Measures.Distinct().ToList();
string findText = null;
string replaceText = null;

if (selectedMeasures.Count == 0)
{
    var pickedMeasures = new HashSet<Measure>();

    using (var picker = new Form())
    {
        picker.Text = "Select measures to rename";
        picker.Size = BigDialogSize();
        picker.MinimumSize = new Size(760, 520);
        picker.StartPosition = FormStartPosition.CenterScreen;
        picker.FormBorderStyle = FormBorderStyle.Sizable;
        picker.MaximizeBox = true;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblSearch = new Label() { Text = "Search table, display folder, or measure:", AutoSize = true, Font = font };
        var txtSearch = new TextBox() { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 4, 0, 8) };
        var tree = new TreeView() { Dock = DockStyle.Fill, Font = font, CheckBoxes = true };
        var lblStatus = new Label() { Text = "0 measures selected", AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisibleButton = new Button() { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearButton = new Button() { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button() { Text = "Continue", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button() { Text = "Cancel", AutoSize = true, Font = font, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange(new Control[] { checkVisibleButton, clearButton, okButton, cancelButton });

        Action updateStatus = () => lblStatus.Text = pickedMeasures.Count + " measure(s) selected";

        bool suppressTreeEvents = false;
        Action<TreeNode, bool> setNodeAndDescendants = null;
        setNodeAndDescendants = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var measure = node.Tag as Measure;
            if (measure != null)
            {
                if (isChecked) pickedMeasures.Add(measure);
                else pickedMeasures.Remove(measure);
            }
            foreach (TreeNode child in node.Nodes) setNodeAndDescendants(child, isChecked);
        };

        Action rebuildTree = () =>
        {
            suppressTreeEvents = true;
            tree.BeginUpdate();
            tree.Nodes.Clear();

            var search = (txtSearch.Text ?? "").Trim();
            var filtered = allMeasures
                .Where(m => string.IsNullOrEmpty(search)
                    || m.Table.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || (m.DisplayFolder ?? "").IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                    || m.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(m => m.Table.Name)
                .ThenBy(m => m.DisplayFolder)
                .ThenBy(m => m.Name)
                .ToList();

            var nodeLookup = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var measure in filtered)
            {
                var pathParts = new List<string>() { measure.Table.Name };
                if (!string.IsNullOrWhiteSpace(measure.DisplayFolder))
                    pathParts.AddRange(measure.DisplayFolder.Split('\\').Where(p => !string.IsNullOrWhiteSpace(p)));

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

                level.Add(new TreeNode(measure.Name) { Tag = measure, Checked = pickedMeasures.Contains(measure) });
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
            foreach (TreeNode node in tree.Nodes) setNodeAndDescendants(node, true);
            suppressTreeEvents = false;
            updateStatus();
        };
        clearButton.Click += (s, e) =>
        {
            pickedMeasures.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (pickedMeasures.Count == 0)
            {
                MessageBox.Show("Select at least one measure.", "No measures selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        picker.Controls.Add(layout);
        rebuildTree();

        if (picker.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }

    selectedMeasures = pickedMeasures.OrderBy(m => m.Table.Name).ThenBy(m => m.Name).ToList();
}

// Suggest the longest contiguous sequence of complete words shared by all
// selected measure names. This is intentionally conservative: a suggestion
// is made only when at least two measures share a complete word/phrase.
string suggestedFindText = "";
if (selectedMeasures.Count >= 2)
{
    var firstWords = selectedMeasures[0].Name
        .Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

    for (int phraseLength = firstWords.Length; phraseLength >= 1 && string.IsNullOrEmpty(suggestedFindText); phraseLength--)
    {
        for (int start = 0; start <= firstWords.Length - phraseLength; start++)
        {
            var candidate = string.Join(" ", firstWords.Skip(start).Take(phraseLength));
            bool presentInAll = selectedMeasures.Skip(1).All(m =>
                (" " + m.Name + " ").IndexOf(" " + candidate + " ", StringComparison.OrdinalIgnoreCase) >= 0);
            if (presentInAll)
            {
                suggestedFindText = candidate;
                break;
            }
        }
    }
}

using (var form = new Form())
{
    form.Text = "Rename " + selectedMeasures.Count + " selected measure(s)";
    form.AutoSize = true;
    form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.MaximizeBox = false;
    form.MinimizeBox = false;
    form.StartPosition = FormStartPosition.CenterScreen;
    var font = new Font("Segoe UI", 9);
    const int textWidth = 620;

    var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 6, Padding = new Padding(12) };
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

    var lblSelection = new Label() {
        Text = selectedMeasures.Count + " measure(s) will be renamed. The same old -> new name mapping will be applied to every PBIR report you select in the next step.",
        AutoSize = true, MaximumSize = new Size(textWidth, 0), Font = font, Margin = new Padding(0, 0, 0, 12)
    };
    var lblFind = new Label() { Text = "Find:", AutoSize = true, Font = font, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 8, 6) };
    var txtFind = new TextBox() { Width = 480, Font = font, Text = suggestedFindText, Margin = new Padding(0, 0, 0, 6) };
    var lblReplace = new Label() { Text = "Replace with:", AutoSize = true, Font = font, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 8, 6) };
    var txtReplace = new TextBox() { Width = 480, Font = font, Margin = new Padding(0, 0, 0, 6) };

    var lblHint = new Label() {
        Text = "Applied to every selected measure's current name (case-sensitive). Measures not containing the Find text are skipped.",
        AutoSize = true, MaximumSize = new Size(textWidth, 0), Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 6)
    };
    var lblSync = new Label() {
        Text = "Important: save the model and deploy/update it in the service so the connected PBIR reports can find the renamed measures. Updating the PBIR files alone is not enough.",
        AutoSize = true, MaximumSize = new Size(textWidth, 0), Font = font, ForeColor = Color.DarkRed, Margin = new Padding(0, 0, 0, 12)
    };
    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    var okButton = new Button() { Text = "Continue", AutoSize = true, DialogResult = DialogResult.OK, Font = font, Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button() { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel, Font = font };
    buttons.Controls.AddRange(new Control[] { okButton, cancelButton });

    layout.Controls.Add(lblSelection, 0, 0);
    layout.SetColumnSpan(lblSelection, 2);
    layout.Controls.Add(lblFind, 0, 1);
    layout.Controls.Add(txtFind, 1, 1);
    layout.Controls.Add(lblReplace, 0, 2);
    layout.Controls.Add(txtReplace, 1, 2);
    layout.Controls.Add(lblHint, 0, 3);
    layout.SetColumnSpan(lblHint, 2);
    layout.Controls.Add(lblSync, 0, 4);
    layout.SetColumnSpan(lblSync, 2);
    layout.Controls.Add(buttons, 0, 5);
    layout.SetColumnSpan(buttons, 2);

    form.Controls.Add(layout);
    form.AcceptButton = okButton;
    form.CancelButton = cancelButton;

    if (form.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }

    findText = txtFind.Text;
    replaceText = txtReplace.Text ?? "";
}

if (selectedMeasures.Count == 0) Error("No measures selected.");
if (string.IsNullOrEmpty(findText)) Error("'Find' text cannot be empty.");

// ---------------------------------------------------------------------------
// Step 2: compute old -> new names, validate (skip on no-match / collision)
// ---------------------------------------------------------------------------
var rows = selectedMeasures.Select(m => new {
    Measure = m,
    Table = m.Table.Name,
    OldName = m.Name,
    NewName = m.Name.Contains(findText) ? m.Name.Replace(findText, replaceText) : m.Name
}).ToList();

// New names that would collide with each other within this batch
var duplicateNames = rows
    .Where(r => r.NewName != r.OldName)
    .GroupBy(r => r.NewName, StringComparer.OrdinalIgnoreCase)
    .Where(g => g.Count() > 1)
    .Select(g => g.Key)
    .ToList();
var dupWithinBatch = new HashSet<string>(duplicateNames, StringComparer.OrdinalIgnoreCase);

var plan = new List<PlanRow>();
foreach (var r in rows)
{
    var pr = new PlanRow { Measure = r.Measure, Table = r.Table, OldName = r.OldName, NewName = r.NewName };
    if (r.NewName == r.OldName)
    {
        pr.Action = "Skip - 'Find' text not present";
    }
    else if (dupWithinBatch.Contains(r.NewName))
    {
        pr.Action = "SKIP - duplicate result within this batch";
    }
    else if (Model.AllMeasures.Any(o => o != r.Measure && string.Equals(o.Name, r.NewName, StringComparison.OrdinalIgnoreCase)))
    {
        pr.Action = "SKIP - name already exists in model";
    }
    else
    {
        pr.Action = "Rename";
    }
    plan.Add(pr);
}

var toRename = plan.Where(p => p.Action == "Rename").ToList();
if (toRename.Count == 0)
{
    var skipped = new System.Data.DataTable();
    skipped.Columns.Add("Table");
    skipped.Columns.Add("Old name");
    skipped.Columns.Add("New name");
    skipped.Columns.Add("Action");
    foreach (var p in plan) skipped.Rows.Add(p.Table, p.OldName, p.NewName, p.Action);
    Output(skipped);
    Info("Nothing to rename - every selected measure was skipped (no match, or name collision). No changes were made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 3: select root folder and PBIR reports
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
    form.Text = "Select PBIR reports to update";
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
        Text = "Selected reports receive the same old -> new measure name mapping. Reports live-connected to the published model only work after the renamed model is deployed/updated in the service.",
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

// Lookup: "Table.OldName" -> PlanRow, used to match queryRef / selector.metadata strings
var qualifiedMap = new Dictionary<string, PlanRow>();
foreach (var p in toRename) qualifiedMap[p.Table + "." + p.OldName] = p;

// Recursively rewrite one parsed JSON document in place. aliasMap resolves
// semantic-query Source aliases (for example "m") through filter.From.
// forcedFilterMeasure is set by a filter's outer field reference, allowing a
// stale nested advanced-filter Property to be corrected to that same measure.
// Returns the number of individual reference edits made.
Func<JToken, string, Dictionary<string, string>, PlanRow, int> walk = null;
walk = (token, parentKey, aliasMap, forcedFilterMeasure) =>
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
                if (!string.IsNullOrEmpty(alias) && !string.IsNullOrEmpty(entity))
                    localAliases[alias] = entity;
            }
        }

        // A filter's outer field is authoritative. If it is being renamed,
        // nested semantic-query expressions in filter.Where must reference
        // the same new measure even when their old Property is already stale.
        PlanRow localForcedFilterMeasure = forcedFilterMeasure;
        var filterMeasure = obj["field"]?["Measure"];
        if (filterMeasure != null && obj["filter"] != null)
        {
            var filterEntity = (string)filterMeasure["Expression"]?["SourceRef"]?["Entity"];
            var filterProperty = (string)filterMeasure["Property"];
            PlanRow filterHit;
            if (filterEntity != null && filterProperty != null
                && qualifiedMap.TryGetValue(filterEntity + "." + filterProperty, out filterHit))
                localForcedFilterMeasure = filterHit;
        }

        // A) Structural field reference: {"Measure": {"Expression":{"SourceRef":{"Entity":"T"}}, "Property":"Old"}}
        //    Handles projections and filter fields. Nested advanced-filter
        //    expressions may use Source aliases instead of Entity names.
        if (parentKey == "Measure")
        {
            var sourceRef = obj["Expression"]?["SourceRef"];
            var entity = (string)sourceRef?["Entity"];
            if (entity == null)
            {
                var sourceAlias = (string)sourceRef?["Source"];
                if (sourceAlias != null) localAliases.TryGetValue(sourceAlias, out entity);
            }
            var propVal = (string)obj["Property"];
            if (entity != null && propVal != null)
            {
                PlanRow hit;
                if (localForcedFilterMeasure != null
                    && string.Equals(localForcedFilterMeasure.Table, entity, StringComparison.OrdinalIgnoreCase))
                {
                    if (propVal != localForcedFilterMeasure.NewName)
                    {
                        obj["Property"] = localForcedFilterMeasure.NewName;
                        changes++;
                    }
                }
                else if (qualifiedMap.TryGetValue(entity + "." + propVal, out hit))
                {
                    obj["Property"] = hit.NewName;
                    changes++;
                }
            }
        }

        // B) queryRef (+ its sibling nativeQueryRef, only when it still equals the default/old name)
        var qrToken = obj["queryRef"];
        if (qrToken != null && qrToken.Type == JTokenType.String)
        {
            PlanRow hit;
            if (qualifiedMap.TryGetValue((string)qrToken, out hit))
            {
                obj["queryRef"] = hit.Table + "." + hit.NewName;
                changes++;
                var nqrToken = obj["nativeQueryRef"];
                if (nqrToken != null && nqrToken.Type == JTokenType.String && (string)nqrToken == hit.OldName)
                {
                    obj["nativeQueryRef"] = hit.NewName;
                    changes++;
                }
            }
        }

        // C) selector.metadata ("Table.OldName" string used by conditional-formatting/entry selectors)
        var metaToken = obj["metadata"];
        if (metaToken != null && metaToken.Type == JTokenType.String)
        {
            PlanRow hit;
            if (qualifiedMap.TryGetValue((string)metaToken, out hit))
            {
                obj["metadata"] = hit.Table + "." + hit.NewName;
                changes++;
            }
        }

        foreach (var prop in obj.Properties().ToList())
            changes += walk(prop.Value, prop.Name, localAliases, localForcedFilterMeasure);
    }
    else
    {
        var arr = token as JArray;
        if (arr != null)
            foreach (var item in arr) changes += walk(item, parentKey, aliasMap, forcedFilterMeasure);
    }
    return changes;
};

var filesToWrite = new List<Tuple<string, string, JObject, int>>(); // (definition folder, path, mutated root, refCount)
if (toRename.Count > 0)
{
    foreach (var reportFolder in reportFolders)
    {
        var definitionFolder = Path.Combine(reportFolder, "definition");
        if (!Directory.Exists(definitionFolder))
        {
            Warning("Skipping selected folder because 'definition' was not found: " + reportFolder);
            continue;
        }

        var candidateFiles = new List<string>();
        var reportJson = Path.Combine(definitionFolder, "report.json");
        if (File.Exists(reportJson)) candidateFiles.Add(reportJson);
        candidateFiles.AddRange(Directory.GetFiles(definitionFolder, "page.json", SearchOption.AllDirectories));
        candidateFiles.AddRange(Directory.GetFiles(definitionFolder, "visual.json", SearchOption.AllDirectories));

        foreach (var file in candidateFiles)
        {
            JObject root;
            try { root = JObject.Parse(File.ReadAllText(file)); } catch { continue; }
            int refCount = walk(root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), null);
            if (refCount > 0) filesToWrite.Add(Tuple.Create(definitionFolder, file, root, refCount));
        }
    }
}

// ---------------------------------------------------------------------------
// Step 4: combined preview + confirmation
// ---------------------------------------------------------------------------
var dt = new System.Data.DataTable();
dt.Columns.Add("Scope");
dt.Columns.Add("Location");
dt.Columns.Add("Old name");
dt.Columns.Add("New name");
dt.Columns.Add("Action");

foreach (var p in plan)
    dt.Rows.Add("Model (TOM)", p.Table, p.OldName, p.NewName, p.Action);

foreach (var f in filesToWrite)
{
    var rel = f.Item2.Substring(f.Item1.Length).TrimStart('\\', '/');
    dt.Rows.Add("PBIR", f.Item1 + " | " + rel, "", "", "Update " + f.Item4 + " reference(s)");
}

if (!ShowGrid(
        "Preview measure rename",
        "Review the model renames and affected PBIR files. Select Continue to confirm and apply the changes.",
        dt,
        "Continue"))
{
    Info("Cancelled - no changes made.");
    return;
}

var confirm = MessageBox.Show(
    "Rename " + toRename.Count + " measure(s) once in the model and update " + filesToWrite.Count + " PBIR file(s) across " + reportFolders.Count + " selected report(s)?\n\n" +
    "PBIR edits cannot be undone in Tabular Editor. If the model/reports are under source control, commit/sync first if required.\n\n" +
    "After running:\n" +
    "1. First save/deploy/sync the semantic-model changes. Updating PBIR alone is not sufficient.\n" +
    "2. Then apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    "Confirm rename", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no changes made."); return; }

// ---------------------------------------------------------------------------
// Step 5: apply - model renames first, then write the mutated PBIR files
// ---------------------------------------------------------------------------
int modelRenamesApplied = 0;
foreach (var p in toRename)
{
    p.Measure.Name = p.NewName;
    modelRenamesApplied++;
}

int filesWritten = 0;
int referencesWritten = 0;
foreach (var f in filesToWrite)
{
    try
    {
        File.WriteAllText(f.Item2, f.Item3.ToString(Newtonsoft.Json.Formatting.Indented));
        filesWritten++;
        referencesWritten += f.Item4;
    }
    catch (Exception ex)
    {
        Error(string.Format(
            "The update stopped while writing:\n{0}\n\n" +
            "Changes already made before the error:\n" +
            "- {1} measure rename(s) in the open model\n" +
            "- {2} PBIR reference update(s) across {3} file(s)\n\n" +
            "Error: {4}",
            f.Item2,
            modelRenamesApplied,
            referencesWritten,
            filesWritten,
            ex.Message));
    }
}

Info(string.Format(
    "Renamed {0} measure(s) in the model and updated {1} reference(s) across {2} PBIR file(s).\n\n" +
    "Next:\n" +
    "1. First save/deploy/sync the semantic-model changes.\n" +
    "2. Then apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.\n\n" +
    "The model rename was performed once; additional selected reports only received PBIR reference updates.",
    modelRenamesApplied,
    referencesWritten,
    filesWritten));
