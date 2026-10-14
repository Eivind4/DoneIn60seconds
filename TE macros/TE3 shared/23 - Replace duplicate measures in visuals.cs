// ============================================================================
// 17 - Replace field in visuals (PBIR).cs
//
// PBIR-only macro that replaces ONE measure or column with another measure or
// column (same type) in report visuals. The semantic model is not changed.
//
// Workflow:
//   1. Field to replace: uses the single measure/column selected in Tabular
//      Editor; otherwise pick it from a searchable tree
//        Table > Display folder > Measures / Columns > Field
//   2. Replacement field: picked from the same kind of tree (same type only).
//   3. Pick a folder; every definition.pbir below it is found and you select
//      which reports to include.
//   4. Visuals that reference the field are listed. Apply to all of them, or
//      select visuals:  Report > Page > Visual type > Visual
//   5. Related measures (only asked when found in the chosen visuals):
//        - Tooltips          measures in the visual's Tooltips well
//        - Cond. formatting  measures used in formatting tied to the replaced
//                            field (selector.metadata = the field)
//        - Other formatting  measures used elsewhere in the visual formatting
//                            (dynamic titles, reference lines, ...)
//      Per related measure choose Keep, Replace (with a model measure) or
//      Remove. A replacement is suggested when a measure exists with the
//      replaced field's name swapped for the new one (e.g. "Sales Color" ->
//      "Sales PY Color").
//   6. Display names (only asked when the replaced field has explicit display
//      names): set the new display name per existing value. Blank or equal to
//      the new field name removes the display name (shows the field name).
//   7. Preview, confirm, write.
//
// What is replaced in the chosen visuals (visual.json only):
//   - Field references in every well (projections), sort, visual filters
//     (incl. advanced filter From/Where aliases), and formatting expressions.
//   - queryRef (also inside aggregations such as Sum(Table.Column)),
//     nativeQueryRef when it still equals the old field name, and
//     selector.metadata, so conditional formatting stays attached.
//
// Not changed (reported in the preview where detected):
//   - Page and report filters (page.json / report.json) - they are not visual
//     specific. Use macro 15 to replace a field across entire reports.
//   - DAX text: visual calculations and report-level measures
//     (reportExtensions.json) that mention the field by name.
//   - If the new field is already in the visual, it will appear twice.
//
// IMPORTANT:
//   - Only PBIR files are changed; the semantic model is not affected. PBIR
//     edits are not undoable through Tabular Editor's undo stack. If the
//     reports are under source control, commit/sync first if required.
//   - Afterwards, apply external changes using Desktop Bridge, or close and
//     re-open the PBIR file in Power BI Desktop to see the changes.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
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

class FieldRef
{
    public string Kind;          // "Measure" or "Column"
    public string Table;
    public string Name;
    public string DisplayFolder;
    public string Key { get { return Kind + "|" + Table + "|" + Name; } }
    public string QueryRef { get { return Table + "." + Name; } }
    public string Label { get { return Table + "[" + Name + "]"; } }
}

class VisualDoc
{
    public string DefinitionFolder;
    public string File;
    public JObject Root;
    public string ReportName;
    public string PageDisplay;
    public string VisualType;
    public string VisualLabel;
    public int Changes;
    public List<string> Notes = new List<string>();
    public string FullLabel { get { return ReportName + " | " + PageDisplay + " | " + VisualLabel; } }
}

// One place where a related (non-replaced) measure is used in a visual
class RelatedUse
{
    public VisualDoc Visual;
    public string Usage;              // "Tooltip", "Conditional formatting", "Other formatting"
    public string Detail;             // e.g. values.fontColor
    public string MeasureTable;
    public string MeasureName;
    // Tooltip
    public JArray ProjectionArray;
    public JObject Projection;
    // Formatting
    public JObject Container;         // visual.objects or visual.visualContainerObjects
    public string ObjectName;
    public JObject Entry;             // array element { properties, selector }
    public JObject Properties;
    public string PropertyName;
    public List<JObject> MeasureNodes = new List<JObject>();
}

class RelatedGroup
{
    public string Usage;
    public string MeasureTable;
    public string MeasureName;
    public List<RelatedUse> Uses = new List<RelatedUse>();
    public string Suggestion;
    public string Action = "Keep";    // Keep / Replace / Remove
    public Measure ReplaceWith;
}

class DisplayNameUse
{
    public VisualDoc Visual;
    public JObject Projection;
    public string OldDisplayName;
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------
Func<string, string, bool> SameName = (a, b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

Func<double, double, Size> DialogSize = (w, h) =>
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * w), (int)(wa.Height * h));
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

Func<string, string> LiteralText = raw =>
{
    if (string.IsNullOrEmpty(raw)) return null;
    return raw.Length >= 2 && raw.StartsWith("'") && raw.EndsWith("'")
        ? raw.Substring(1, raw.Length - 2).Replace("''", "'")
        : raw;
};

// Resizable dialog with a TableLayoutPanel; rows are added by the caller.
Func<Form, string, int, Size, TableLayoutPanel> CreateDialogLayout = (dlg, title, rowCount, size) =>
{
    dlg.Text = title;
    dlg.Size = size;
    dlg.MinimumSize = new Size(640, 420);
    dlg.StartPosition = FormStartPosition.CenterScreen;
    dlg.FormBorderStyle = FormBorderStyle.Sizable;
    dlg.MaximizeBox = true;
    dlg.MinimizeBox = false;
    dlg.Font = new Font("Segoe UI", 9);
    var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = rowCount, ColumnCount = 1, Padding = new Padding(12) };
    dlg.Controls.Add(panel);
    return panel;
};

Func<string, Label> HeadingLabel = text => new Label {
    Text = text, AutoSize = true, MaximumSize = new Size(1100, 0),
    Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8)
};

// Read-only grid preview with a heading and Continue/Cancel buttons.
Func<string, string, System.Data.DataTable, string, bool> ShowGrid = (title, heading, table, continueText) =>
{
    using (var dlg = new Form())
    {
        var panel = CreateDialogLayout(dlg, title, 3, DialogSize(0.85, 0.85));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var previewGrid = new DataGridView {
            Dock = DockStyle.Fill, ReadOnly = true,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            BackgroundColor = SystemColors.Window, DataSource = table
        };
        var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var okButton = new Button { Text = continueText, AutoSize = true, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttonRow.Controls.AddRange(new Control[] { okButton, cancelButton });
        panel.Controls.Add(HeadingLabel(heading), 0, 0);
        panel.Controls.Add(previewGrid, 0, 1);
        panel.Controls.Add(buttonRow, 0, 2);
        dlg.AcceptButton = okButton;
        dlg.CancelButton = cancelButton;
        return dlg.ShowDialog() == DialogResult.OK;
    }
};

// Highlight style for editable grid columns (same as macros 12, 13 and 15)
var EditBack = Color.FromArgb(255, 244, 179);
var EditHeaderBack = Color.FromArgb(255, 214, 102);
Action<DataGridViewColumn> HighlightColumn = column =>
{
    column.DefaultCellStyle = new DataGridViewCellStyle {
        BackColor = EditBack, ForeColor = Color.Black, Font = new Font("Segoe UI", 9, FontStyle.Bold),
        SelectionBackColor = EditHeaderBack, SelectionForeColor = Color.Black
    };
    column.HeaderCell.Style.BackColor = EditHeaderBack;
    column.HeaderCell.Style.ForeColor = Color.Black;
};
Func<DataGridView> CreateEditGrid = () =>
{
    var editGrid = new DataGridView {
        Dock = DockStyle.Fill,
        AutoGenerateColumns = false,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.CellSelect,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
        BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false,   // required for custom header colours
        EditMode = DataGridViewEditMode.EditOnEnter
    };
    editGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
    editGrid.DefaultCellStyle.ForeColor = Color.FromArgb(90, 90, 90);
    // Commit combo/check cells immediately so dependent columns update
    editGrid.CurrentCellDirtyStateChanged += (s, e) =>
    {
        if (editGrid.IsCurrentCellDirty && (editGrid.CurrentCell is DataGridViewComboBoxCell || editGrid.CurrentCell is DataGridViewCheckBoxCell))
            editGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
    };
    return editGrid;
};

// Single-select field picker: Table > Display folder > (Measures / Columns) > Field
Func<string, string, string, FieldRef, FieldRef> PickField = (title, heading, kindFilter, exclude) =>
{
    FieldRef picked = null;
    var choices = new List<FieldRef>();
    if (kindFilter == null || kindFilter == "Measure")
        choices.AddRange(Model.AllMeasures.Select(ms => new FieldRef { Kind = "Measure", Table = ms.Table.Name, Name = ms.Name, DisplayFolder = ms.DisplayFolder ?? "" }));
    if (kindFilter == null || kindFilter == "Column")
        choices.AddRange(Model.AllColumns.Select(cl => new FieldRef { Kind = "Column", Table = cl.Table.Name, Name = cl.Name, DisplayFolder = cl.DisplayFolder ?? "" }));
    if (exclude != null) choices = choices.Where(ch => ch.Key != exclude.Key).ToList();

    using (var dlg = new Form())
    {
        var panel = CreateDialogLayout(dlg, title, 5, DialogSize(0.5, 0.8));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var searchBox = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 8) };
        var fieldTree = new TreeView { Dock = DockStyle.Fill, HideSelection = false };
        var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttonRow.Controls.AddRange(new Control[] { okButton, cancelButton });

        Action rebuild = () =>
        {
            fieldTree.BeginUpdate();
            fieldTree.Nodes.Clear();
            var search = (searchBox.Text ?? "").Trim();
            var filtered = choices.Where(ch =>
                string.IsNullOrEmpty(search)
                || ch.Table.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || ch.DisplayFolder.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || ch.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(ch => ch.Table).ThenBy(ch => ch.Kind).ThenBy(ch => ch.DisplayFolder).ThenBy(ch => ch.Name).ToList();

            var lookup = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var ch in filtered)
            {
                var parts = new List<string> { ch.Table };
                if (kindFilter == null) parts.Add(ch.Kind == "Measure" ? "Measures" : "Columns");
                if (!string.IsNullOrWhiteSpace(ch.DisplayFolder))
                    parts.AddRange(ch.DisplayFolder.Split('\\').Where(p => !string.IsNullOrWhiteSpace(p)));
                TreeNodeCollection level = fieldTree.Nodes;
                string path = "";
                foreach (var part in parts)
                {
                    path = path.Length == 0 ? part : path + "\\" + part;
                    TreeNode folderNode;
                    if (!lookup.TryGetValue(path, out folderNode))
                    {
                        folderNode = new TreeNode(part);
                        level.Add(folderNode);
                        lookup[path] = folderNode;
                    }
                    level = folderNode.Nodes;
                }
                level.Add(new TreeNode(ch.Name) { Tag = ch });
            }
            fieldTree.ExpandAll();
            if (fieldTree.Nodes.Count > 0) fieldTree.Nodes[0].EnsureVisible();
            fieldTree.EndUpdate();
        };

        Action accept = () =>
        {
            var choice = fieldTree.SelectedNode == null ? null : fieldTree.SelectedNode.Tag as FieldRef;
            if (choice == null)
            {
                MessageBox.Show("Select a field (a leaf in the tree).", "No field selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            picked = choice;
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };
        searchBox.TextChanged += (s, e) => rebuild();
        okButton.Click += (s, e) => accept();
        fieldTree.NodeMouseDoubleClick += (s, e) => { if (e.Node.Tag is FieldRef) accept(); };

        panel.Controls.Add(HeadingLabel(heading), 0, 0);
        panel.Controls.Add(new Label { Text = "Search table, display folder, or field (double-click a field to select it):", AutoSize = true }, 0, 1);
        panel.Controls.Add(searchBox, 0, 2);
        panel.Controls.Add(fieldTree, 0, 3);
        panel.Controls.Add(buttonRow, 0, 4);
        dlg.AcceptButton = okButton;
        dlg.CancelButton = cancelButton;
        dlg.Shown += (s, e) => searchBox.Focus();
        rebuild();
        return dlg.ShowDialog() == DialogResult.OK ? picked : null;
    }
};

// Replaces "Table.Field" as a whole value or inside an aggregation "Sum(Table.Field)"
Func<string, FieldRef, FieldRef, string> ReplaceRefToken = (value, fromField, toField) =>
{
    if (value == null) return null;
    if (SameName(value, fromField.QueryRef)) return toField.QueryRef;
    return Regex.Replace(value, @"(?<=\()" + Regex.Escape(fromField.QueryRef) + @"(?=\))",
        toField.QueryRef.Replace("$", "$$"), RegexOptions.IgnoreCase);
};

Func<JToken, string, string, bool> IsMeasureNode = (node, table, name) =>
    node is JObject
    && SameName((string)node["Expression"]?["SourceRef"]?["Entity"], table)
    && SameName((string)node["Property"], name);

// ---------------------------------------------------------------------------
// Step 1: field to replace and replacement field
// ---------------------------------------------------------------------------
FieldRef sourceField = null;
var tomSelection = Selected.Measures.Select(ms => new FieldRef { Kind = "Measure", Table = ms.Table.Name, Name = ms.Name, DisplayFolder = ms.DisplayFolder ?? "" })
    .Concat(Selected.Columns.Select(cl => new FieldRef { Kind = "Column", Table = cl.Table.Name, Name = cl.Name, DisplayFolder = cl.DisplayFolder ?? "" }))
    .ToList();
if (tomSelection.Count == 1)
    sourceField = tomSelection[0];
else
{
    sourceField = PickField("Field to replace", "Select the measure or column to replace in visuals" +
        (tomSelection.Count > 1 ? " (more than one object is selected in Tabular Editor)" : "") + ":", null, null);
    if (sourceField == null) { Info("Cancelled - no changes made."); return; }
}

var targetField = PickField("Replacement " + sourceField.Kind.ToLower(),
    "Replace " + sourceField.Kind.ToLower() + " " + sourceField.Label + " with:", sourceField.Kind, sourceField);
if (targetField == null) { Info("Cancelled - no changes made."); return; }

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
    .Select(fp => new ReportChoice {
        DefinitionFile = fp,
        ReportFolder = Path.GetDirectoryName(fp),
        ReportName = CleanReportName(Path.GetDirectoryName(fp)),
        ConnectionSummary = ReadConnectionSummary(fp),
        Selected = true
    })
    .OrderBy(rc => rc.ReportName).ThenBy(rc => rc.ReportFolder)
    .ToList();

if (reportChoices.Count == 0)
{
    Info("No definition.pbir files were found below:\n" + rootFolder + "\n\nNo changes made.");
    return;
}

using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Select PBIR reports", 4, DialogSize(0.85, 0.85));
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
        foreach (var rc in reportChoices.Where(rc => string.IsNullOrEmpty(searchBox.Text) || rc.Label.IndexOf(searchBox.Text, StringComparison.OrdinalIgnoreCase) >= 0))
            rc.Selected = true;
        rebuild();
    };
    clearAll.Click += (s, e) => { foreach (var rc in reportChoices) rc.Selected = false; rebuild(); };
    okButton.Click += (s, e) =>
    {
        if (!reportChoices.Any(rc => rc.Selected))
        {
            MessageBox.Show("Select at least one report.", "No reports selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(HeadingLabel("Replace " + sourceField.Label + " with " + targetField.Label + " - select the reports to search:"), 0, 0);
    panel.Controls.Add(searchBox, 0, 1);
    panel.Controls.Add(reportList, 0, 2);
    panel.Controls.Add(buttonRow, 0, 3);
    dlg.AcceptButton = okButton;
    dlg.CancelButton = cancelButton;
    rebuild();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

var reportFolders = reportChoices.Where(rc => rc.Selected).Select(rc => rc.ReportFolder)
    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

// ---------------------------------------------------------------------------
// Replacement engine for the source field (ported from macro 15, single field)
// ---------------------------------------------------------------------------
Func<JToken, string, Dictionary<string, string>, bool, int> MutateSource = null;
MutateSource = (token, parentKey, aliasMap, forced) =>
{
    int changes = 0;
    var obj = token as JObject;
    if (obj != null)
    {
        var localAliases = new Dictionary<string, string>(aliasMap, StringComparer.OrdinalIgnoreCase);

        // A filter's outer field is authoritative: nested advanced-filter
        // expressions (From/Where aliases) follow it to the new field.
        bool localForced = forced;
        var outerField = obj["field"]?[sourceField.Kind];
        if (outerField != null && obj["filter"] != null
            && SameName((string)outerField["Expression"]?["SourceRef"]?["Entity"], sourceField.Table)
            && SameName((string)outerField["Property"], sourceField.Name))
            localForced = true;

        var fromArray = obj["From"] as JArray;
        if (fromArray != null)
        {
            foreach (var from in fromArray.OfType<JObject>())
            {
                var alias = (string)from["Name"];
                var entity = (string)from["Entity"];
                if (string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(entity)) continue;
                if (localForced && SameName(entity, sourceField.Table) && !SameName(entity, targetField.Table))
                {
                    from["Entity"] = targetField.Table;
                    entity = targetField.Table;
                    changes++;
                }
                localAliases[alias] = entity;
            }
        }

        if (parentKey == sourceField.Kind)
        {
            var sourceRef = obj["Expression"]?["SourceRef"];
            var entityToken = sourceRef?["Entity"];
            var table = (string)entityToken;
            var alias = (string)sourceRef?["Source"];
            if (table == null && alias != null) localAliases.TryGetValue(alias, out table);
            var field = (string)obj["Property"];

            bool isSource = SameName(table, sourceField.Table) && SameName(field, sourceField.Name);
            // Alias entity already switched to the target table in a forced filter
            if (!isSource && localForced && alias != null && SameName(table, targetField.Table) && SameName(field, sourceField.Name))
                isSource = true;

            if (isSource)
            {
                if (entityToken != null && !SameName(table, targetField.Table))
                {
                    sourceRef["Entity"] = targetField.Table;
                    changes++;
                }
                if (!string.Equals(field, targetField.Name, StringComparison.Ordinal))
                {
                    obj["Property"] = targetField.Name;
                    changes++;
                }
            }
        }

        var queryRef = obj["queryRef"];
        if (queryRef != null && queryRef.Type == JTokenType.String)
        {
            var oldValue = (string)queryRef;
            var newValue = ReplaceRefToken(oldValue, sourceField, targetField);
            if (newValue != oldValue)
            {
                obj["queryRef"] = newValue;
                changes++;
                var native = obj["nativeQueryRef"];
                if (native != null && native.Type == JTokenType.String)
                {
                    var nativeValue = (string)native;
                    if (string.Equals(nativeValue, sourceField.Name, StringComparison.Ordinal))
                    {
                        obj["nativeQueryRef"] = targetField.Name;
                        changes++;
                    }
                    else if (nativeValue.EndsWith(" of " + sourceField.Name, StringComparison.Ordinal))
                    {
                        obj["nativeQueryRef"] = nativeValue.Substring(0, nativeValue.Length - sourceField.Name.Length) + targetField.Name;
                        changes++;
                    }
                }
            }
        }

        var metadata = obj["metadata"];
        if (metadata != null && metadata.Type == JTokenType.String)
        {
            var oldValue = (string)metadata;
            var newValue = ReplaceRefToken(oldValue, sourceField, targetField);
            if (newValue != oldValue)
            {
                obj["metadata"] = newValue;
                changes++;
            }
        }

        foreach (var property in obj.Properties().ToList())
            changes += MutateSource(property.Value, property.Name, localAliases, localForced);
    }
    else
    {
        var array = token as JArray;
        if (array != null)
            foreach (var item in array.ToList()) changes += MutateSource(item, parentKey, aliasMap, forced);
    }
    return changes;
};

Func<JToken, bool> ReferencesSource = json =>
    MutateSource(json.DeepClone(), null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), false) > 0;

// ---------------------------------------------------------------------------
// Step 3: find visuals that reference the field
// ---------------------------------------------------------------------------
var visualDocs = new List<VisualDoc>();
int pageLevelHits = 0;
int reportExtensionHits = 0;
var nameToken = "[" + sourceField.Name + "]";

foreach (var reportFolder in reportFolders)
{
    var definitionFolder = Path.Combine(reportFolder, "definition");
    var pagesFolder = Path.Combine(definitionFolder, "pages");
    if (!Directory.Exists(pagesFolder))
    {
        Warning("Skipping report because definition\\pages was not found: " + reportFolder);
        continue;
    }
    var reportName = CleanReportName(reportFolder);

    // Report/page filters and report-level measures are not changed; count them for the preview
    var reportJson = Path.Combine(definitionFolder, "report.json");
    try { if (File.Exists(reportJson) && ReferencesSource(JObject.Parse(File.ReadAllText(reportJson)))) pageLevelHits++; } catch { }
    var extensionsJson = Path.Combine(definitionFolder, "reportExtensions.json");
    try { if (File.Exists(extensionsJson) && File.ReadAllText(extensionsJson).IndexOf(nameToken, StringComparison.OrdinalIgnoreCase) >= 0) reportExtensionHits++; } catch { }

    foreach (var pageFolder in Directory.GetDirectories(pagesFolder))
    {
        var pageFile = Path.Combine(pageFolder, "page.json");
        JObject pageJson = null;
        try { if (File.Exists(pageFile)) pageJson = JObject.Parse(File.ReadAllText(pageFile)); } catch { }
        if (pageJson != null && ReferencesSource(pageJson)) pageLevelHits++;
        var pageName = (string)pageJson?["displayName"] ?? Path.GetFileName(pageFolder);

        var visualsFolder = Path.Combine(pageFolder, "visuals");
        if (!Directory.Exists(visualsFolder)) continue;

        foreach (var file in Directory.GetFiles(visualsFolder, "visual.json", SearchOption.AllDirectories))
        {
            JObject visualJson;
            try { visualJson = JObject.Parse(File.ReadAllText(file)); }
            catch (Exception ex) { Warning("Skipping unreadable visual.json:\n" + file + "\n\n" + ex.Message); continue; }
            if (!ReferencesSource(visualJson)) continue;

            var objectId = (string)visualJson["name"] ?? Path.GetFileName(Path.GetDirectoryName(file));
            var title = LiteralText((string)visualJson["visual"]?["visualContainerObjects"]?["title"]?[0]?["properties"]?["text"]?["expr"]?["Literal"]?["Value"]);
            visualDocs.Add(new VisualDoc {
                DefinitionFolder = definitionFolder,
                File = file,
                Root = visualJson,
                ReportName = reportName,
                PageDisplay = pageName,
                VisualType = (string)visualJson["visual"]?["visualType"] ?? "(unknown visual type)",
                VisualLabel = string.IsNullOrWhiteSpace(title) ? objectId : title
            });
        }
    }
}

if (visualDocs.Count == 0)
{
    Info("No visuals in the selected reports reference " + sourceField.Label + "." +
        (pageLevelHits > 0 ? "\n\nIt is used in " + pageLevelHits + " page/report filter file(s), which this macro does not change (use macro 15)." : "") +
        "\n\nNo changes made.");
    return;
}

// ---------------------------------------------------------------------------
// Step 4: all visuals or selected visuals
// ---------------------------------------------------------------------------
bool selectSpecificVisuals = false;
using (var dlg = new Form())
{
    dlg.Text = "Choose visual scope";
    dlg.AutoSize = true;
    dlg.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
    dlg.MaximizeBox = false;
    dlg.MinimizeBox = false;
    dlg.StartPosition = FormStartPosition.CenterScreen;
    dlg.Font = new Font("Segoe UI", 9);

    var panel = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, Padding = new Padding(12) };
    panel.Controls.Add(HeadingLabel(visualDocs.Count + " visual(s) in " + visualDocs.Select(v => v.ReportName).Distinct().Count() + " report(s) use " + sourceField.Label + "."));
    panel.Controls.Add(new Label {
        Text = "Where should it be replaced with " + targetField.Label + "?",
        AutoSize = true, MaximumSize = new Size(560, 0), Margin = new Padding(0, 0, 0, 12)
    });
    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
    var allButton = new Button { Text = "All " + visualDocs.Count + " visuals", AutoSize = true, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0) };
    var selectButton = new Button { Text = "Select visuals...", AutoSize = true, Padding = new Padding(6, 3, 6, 3), Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, Padding = new Padding(6, 3, 6, 3), DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { allButton, selectButton, cancelButton });
    panel.Controls.Add(buttonRow);
    allButton.Click += (s, e) => { selectSpecificVisuals = false; dlg.DialogResult = DialogResult.OK; dlg.Close(); };
    selectButton.Click += (s, e) => { selectSpecificVisuals = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); };
    dlg.Controls.Add(panel);
    dlg.CancelButton = cancelButton;
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

if (selectSpecificVisuals)
{
    var selectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var dlg = new Form())
    {
        var panel = CreateDialogLayout(dlg, "Select visuals", 5, DialogSize(0.85, 0.85));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var regularFont = new Font("Segoe UI", 9);
        var searchBox = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 8) };
        // Bold tree font so bold report names are not clipped; children use the regular font
        var visualTree = new TreeView { Dock = DockStyle.Fill, Font = new Font(regularFont, FontStyle.Bold), CheckBoxes = true };
        var status = new Label { Text = "0 visuals selected", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
        var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var checkVisible = new Button { Text = "Check all visible", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        var clearAll = new Button { Text = "Clear all", AutoSize = true, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Continue", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttonRow.Controls.AddRange(new Control[] { checkVisible, clearAll, okButton, cancelButton });

        bool suppress = false;
        var expandedReports = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Action updateStatus = () => status.Text = selectedFiles.Count + " visual(s) selected";

        Action<TreeNode, bool> setChecked = null;
        setChecked = (node, isChecked) =>
        {
            node.Checked = isChecked;
            var doc = node.Tag as VisualDoc;
            if (doc != null) { if (isChecked) selectedFiles.Add(doc.File); else selectedFiles.Remove(doc.File); }
            foreach (TreeNode child in node.Nodes) setChecked(child, isChecked);
        };
        Func<TreeNode, bool> refreshParents = null;
        refreshParents = node =>
        {
            if (node.Nodes.Count == 0) return node.Checked;
            bool allChecked = true;
            foreach (TreeNode child in node.Nodes) if (!refreshParents(child)) allChecked = false;
            node.Checked = allChecked;
            return allChecked;
        };

        Action rebuild = () =>
        {
            suppress = true;
            visualTree.BeginUpdate();
            visualTree.Nodes.Clear();
            var search = (searchBox.Text ?? "").Trim();
            var filtered = visualDocs.Where(v =>
                search.Length == 0
                || v.ReportName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.PageDisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualType.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                || v.VisualLabel.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(v => v.ReportName).ThenBy(v => v.PageDisplay).ThenBy(v => v.VisualType).ThenBy(v => v.VisualLabel).ToList();

            foreach (var reportGroup in filtered.GroupBy(v => v.ReportName))
            {
                var reportNode = new TreeNode(reportGroup.Key + "  (" + reportGroup.Count() + " visual(s))") { Name = reportGroup.Key };
                visualTree.Nodes.Add(reportNode);
                foreach (var pageGroup in reportGroup.GroupBy(v => v.PageDisplay))
                {
                    var pageNode = new TreeNode(pageGroup.Key) { NodeFont = regularFont };
                    reportNode.Nodes.Add(pageNode);
                    foreach (var typeGroup in pageGroup.GroupBy(v => v.VisualType))
                    {
                        var typeNode = new TreeNode(typeGroup.Key) { NodeFont = regularFont };
                        pageNode.Nodes.Add(typeNode);
                        foreach (var doc in typeGroup)
                            typeNode.Nodes.Add(new TreeNode(doc.VisualLabel) { Tag = doc, NodeFont = regularFont, Checked = selectedFiles.Contains(doc.File) });
                    }
                }
                refreshParents(reportNode);
                reportNode.ExpandAll();
                if (search.Length == 0 && !expandedReports.Contains(reportNode.Name)) reportNode.Collapse(true);
            }
            visualTree.EndUpdate();
            suppress = false;
            updateStatus();
        };

        visualTree.AfterExpand += (s, e) => { if (!suppress && e.Node.Level == 0) expandedReports.Add(e.Node.Name); };
        visualTree.AfterCollapse += (s, e) => { if (!suppress && e.Node.Level == 0) expandedReports.Remove(e.Node.Name); };
        visualTree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;
            setChecked(e.Node, e.Node.Checked);
            foreach (TreeNode reportNode in visualTree.Nodes) refreshParents(reportNode);
            suppress = false;
            updateStatus();
        };
        searchBox.TextChanged += (s, e) => rebuild();
        checkVisible.Click += (s, e) =>
        {
            suppress = true;
            foreach (TreeNode node in visualTree.Nodes) setChecked(node, true);
            suppress = false;
            updateStatus();
        };
        clearAll.Click += (s, e) => { selectedFiles.Clear(); rebuild(); };
        okButton.Click += (s, e) =>
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("Select at least one visual.", "No visuals selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };

        panel.Controls.Add(HeadingLabel("Select the visuals where " + sourceField.Label + " is replaced with " + targetField.Label + ":"), 0, 0);
        panel.Controls.Add(searchBox, 0, 1);
        panel.Controls.Add(visualTree, 0, 2);
        panel.Controls.Add(status, 0, 3);
        panel.Controls.Add(buttonRow, 0, 4);
        dlg.AcceptButton = okButton;
        dlg.CancelButton = cancelButton;
        rebuild();
        if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }
    visualDocs = visualDocs.Where(v => selectedFiles.Contains(v.File)).ToList();
}

// ---------------------------------------------------------------------------
// Step 5: analyse chosen visuals - related measures, display names, notes
// ---------------------------------------------------------------------------
var relatedUses = new List<RelatedUse>();
var displayNameUses = new List<DisplayNameUse>();
var sourceNodeKey = sourceField.Kind;

foreach (var doc in visualDocs)
{
    var visualNode = doc.Root["visual"] as JObject;
    var queryState = visualNode?["query"]?["queryState"] as JObject;

    if (queryState != null)
    {
        foreach (var role in queryState.Properties())
        {
            var projections = role.Value?["projections"] as JArray;
            if (projections == null) continue;
            foreach (var projection in projections.OfType<JObject>())
            {
                // Direct projection of the source: display name + duplicate check
                // Columns may be wrapped in an aggregation: field.Aggregation.Expression.Column
                var sourceNode = projection["field"]?[sourceNodeKey]
                    ?? projection["field"]?["Aggregation"]?["Expression"]?[sourceNodeKey];
                if (sourceNode != null
                    && SameName((string)sourceNode["Expression"]?["SourceRef"]?["Entity"], sourceField.Table)
                    && SameName((string)sourceNode["Property"], sourceField.Name))
                {
                    var displayName = (string)projection["displayName"];
                    if (!string.IsNullOrWhiteSpace(displayName))
                        displayNameUses.Add(new DisplayNameUse { Visual = doc, Projection = projection, OldDisplayName = displayName });
                    if (projections.OfType<JObject>().Any(other =>
                        other != projection
                        && SameName((string)other["field"]?[targetField.Kind]?["Expression"]?["SourceRef"]?["Entity"], targetField.Table)
                        && SameName((string)other["field"]?[targetField.Kind]?["Property"], targetField.Name)))
                        doc.Notes.Add(targetField.Label + " is already in the '" + role.Name + "' well - it will appear twice");
                    continue;
                }

                // Tooltip measures
                if (role.Name == "Tooltips")
                {
                    var tooltipMeasure = projection["field"]?["Measure"];
                    if (tooltipMeasure != null)
                        relatedUses.Add(new RelatedUse {
                            Visual = doc, Usage = "Tooltip", Detail = "Tooltips well",
                            MeasureTable = (string)tooltipMeasure["Expression"]?["SourceRef"]?["Entity"],
                            MeasureName = (string)tooltipMeasure["Property"],
                            ProjectionArray = projections, Projection = projection
                        });
                }
            }
        }
    }

    // Measures used in formatting (conditional formatting, dynamic titles, reference lines, ...)
    foreach (var containerName in new[] { "objects", "visualContainerObjects" })
    {
        var container = visualNode?[containerName] as JObject;
        if (container == null) continue;
        foreach (var objectProperty in container.Properties())
        {
            var entries = objectProperty.Value as JArray;
            if (entries == null) continue;
            foreach (var entry in entries.OfType<JObject>())
            {
                var properties = entry["properties"] as JObject;
                if (properties == null) continue;
                var selectorMetadata = (string)entry["selector"]?["metadata"];
                bool tiedToSource = selectorMetadata != null && ReplaceRefToken(selectorMetadata, sourceField, targetField) != selectorMetadata;

                foreach (var formatProperty in properties.Properties())
                {
                    var measureNodes = formatProperty.Value.SelectTokens("$..Measure").OfType<JObject>()
                        .Where(n => (string)n["Property"] != null
                            && !(IsMeasureNode(n, sourceField.Table, sourceField.Name) && sourceField.Kind == "Measure"))
                        .ToList();
                    foreach (var measureGroup in measureNodes.GroupBy(n => (string)n["Expression"]?["SourceRef"]?["Entity"] + "|" + (string)n["Property"], StringComparer.OrdinalIgnoreCase))
                    {
                        var firstNode = measureGroup.First();
                        var use = new RelatedUse {
                            Visual = doc,
                            Usage = tiedToSource ? "Conditional formatting" : "Other formatting",
                            Detail = objectProperty.Name + "." + formatProperty.Name + (containerName == "visualContainerObjects" ? " (visual container)" : ""),
                            MeasureTable = (string)firstNode["Expression"]?["SourceRef"]?["Entity"],
                            MeasureName = (string)firstNode["Property"],
                            Container = container, ObjectName = objectProperty.Name, Entry = entry,
                            Properties = properties, PropertyName = formatProperty.Name
                        };
                        use.MeasureNodes.AddRange(measureGroup);
                        relatedUses.Add(use);
                    }
                }
            }
        }
    }

    // DAX text that mentions the field by name (visual calculations) cannot be replaced safely
    foreach (var calc in doc.Root.SelectTokens("$..NativeVisualCalculation").OfType<JObject>())
    {
        var expression = (string)calc["Expression"];
        if (expression != null && expression.IndexOf(nameToken, StringComparison.OrdinalIgnoreCase) >= 0)
            doc.Notes.Add("Visual calculation '" + ((string)calc["Name"] ?? "") + "' mentions " + nameToken + " in DAX - update manually");
    }
}

// Suggest a replacement: related measure name with the source name swapped for the target name
Func<string, Measure> FindMeasure = measureName =>
{
    if (string.IsNullOrWhiteSpace(measureName)) return null;
    var cleaned = measureName.Trim();
    var bracket = cleaned.IndexOf('[');
    if (bracket >= 0 && cleaned.EndsWith("]")) cleaned = cleaned.Substring(bracket + 1, cleaned.Length - bracket - 2);
    return Model.AllMeasures.FirstOrDefault(ms => SameName(ms.Name, cleaned));
};

var relatedGroups = relatedUses
    .GroupBy(u => u.Usage + "|" + u.MeasureTable + "|" + u.MeasureName, StringComparer.OrdinalIgnoreCase)
    .Select(g =>
    {
        var group = new RelatedGroup { Usage = g.First().Usage, MeasureTable = g.First().MeasureTable, MeasureName = g.First().MeasureName };
        group.Uses.AddRange(g);
        if (group.MeasureName.IndexOf(sourceField.Name, StringComparison.OrdinalIgnoreCase) >= 0)
        {
            var swapped = Regex.Replace(group.MeasureName, Regex.Escape(sourceField.Name), targetField.Name.Replace("$", "$$"), RegexOptions.IgnoreCase);
            var suggested = FindMeasure(swapped);
            if (suggested != null && !SameName(suggested.Name, group.MeasureName)) group.Suggestion = suggested.Name;
        }
        return group;
    })
    .OrderBy(g => g.Usage == "Tooltip" ? 0 : g.Usage == "Conditional formatting" ? 1 : 2)
    .ThenBy(g => g.MeasureName)
    .ToList();

// ---------------------------------------------------------------------------
// Step 6: related measures - Keep / Replace / Remove
// ---------------------------------------------------------------------------
if (relatedGroups.Count > 0)
{
    using (var dlg = new Form())
    {
        var panel = CreateDialogLayout(dlg, "Related measures in the chosen visuals", 4, DialogSize(0.85, 0.7));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var relatedGrid = CreateEditGrid();
        var actionColumn = new DataGridViewComboBoxColumn { Name = "Action", HeaderText = "Action (edit)", Width = 110, FlatStyle = FlatStyle.Flat };
        actionColumn.Items.AddRange("Keep", "Replace", "Remove");
        HighlightColumn(actionColumn);
        relatedGrid.Columns.Add(actionColumn);
        relatedGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Usage", HeaderText = "Usage", Width = 160, ReadOnly = true });
        relatedGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Measure", HeaderText = "Related measure", Width = 240, ReadOnly = true });
        var replaceColumn = new DataGridViewTextBoxColumn { Name = "ReplaceWith", HeaderText = "Replace with (edit)", Width = 240 };
        HighlightColumn(replaceColumn);
        relatedGrid.Columns.Add(replaceColumn);
        relatedGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Result", HeaderText = "Result", Width = 230, ReadOnly = true });
        relatedGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Uses", HeaderText = "Uses", Width = 50, ReadOnly = true });
        relatedGrid.Columns.Add(new DataGridViewTextBoxColumn {
            Name = "Where", HeaderText = "Where", ReadOnly = true,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 160,
            DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
        });

        foreach (var group in relatedGroups)
        {
            // Default: replace tooltips/conditional formatting when a matching measure exists
            var defaultAction = group.Suggestion != null && group.Usage != "Other formatting" ? "Replace" : "Keep";
            int rowIndex = relatedGrid.Rows.Add(
                defaultAction,
                group.Usage == "Conditional formatting" ? "Cond. formatting on " + sourceField.Name : group.Usage,
                group.MeasureTable + "[" + group.MeasureName + "]",
                group.Suggestion ?? "",
                "",
                group.Uses.Count,
                string.Join(", ", group.Uses.Select(u => u.Detail).Distinct()) + "  |  " + group.Uses.Select(u => u.Visual.File).Distinct().Count() + " visual(s)");
            relatedGrid.Rows[rowIndex].Tag = group;
        }

        Action<int> updateResult = rowIndex =>
        {
            var gridRow = relatedGrid.Rows[rowIndex];
            var action = Convert.ToString(gridRow.Cells["Action"].Value);
            var cell = gridRow.Cells["Result"];
            if (action == "Remove") { cell.Value = "Remove from these visuals"; cell.Style.ForeColor = Color.DarkRed; return; }
            if (action != "Replace") { cell.Value = "Keep as is"; cell.Style.ForeColor = Color.Gray; return; }
            var replacement = FindMeasure(Convert.ToString(gridRow.Cells["ReplaceWith"].Value));
            if (replacement == null) { cell.Value = "Enter an existing measure name"; cell.Style.ForeColor = Color.DarkRed; return; }
            cell.Value = "Replace with " + replacement.Table.Name + "[" + replacement.Name + "]";
            cell.Style.ForeColor = Color.DarkGreen;
        };
        for (int i = 0; i < relatedGrid.Rows.Count; i++) updateResult(i);
        relatedGrid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) updateResult(e.RowIndex); };

        // Measure name suggestions while typing, and typing a name switches the row to Replace
        relatedGrid.EditingControlShowing += (s, e) =>
        {
            var textBox = e.Control as TextBox;
            if (textBox == null) return;
            if (relatedGrid.CurrentCell.OwningColumn.Name != "ReplaceWith") { textBox.AutoCompleteMode = AutoCompleteMode.None; return; }
            var names = new AutoCompleteStringCollection();
            names.AddRange(Model.AllMeasures.Select(ms => ms.Name).Distinct().ToArray());
            textBox.AutoCompleteCustomSource = names;
            textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        };
        relatedGrid.CellEndEdit += (s, e) =>
        {
            if (relatedGrid.Columns[e.ColumnIndex].Name != "ReplaceWith") return;
            var gridRow = relatedGrid.Rows[e.RowIndex];
            if (!string.IsNullOrWhiteSpace(Convert.ToString(gridRow.Cells["ReplaceWith"].Value))
                && Convert.ToString(gridRow.Cells["Action"].Value) == "Keep")
                gridRow.Cells["Action"].Value = "Replace";
        };

        var hint = new Label {
            Text = "Keep = leave the related measure as it is. Replace = swap it for another model measure in the chosen visuals. Remove = take it out of the tooltips / formatting. " +
                   "Suggestions use the related measure's name with '" + sourceField.Name + "' swapped for '" + targetField.Name + "'.",
            AutoSize = true, MaximumSize = new Size(1100, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
        };
        var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var applyButton = new Button { Text = "Apply", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
        var keepAllButton = new Button { Text = "Keep all as is", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttonRow.Controls.AddRange(new Control[] { applyButton, keepAllButton, cancelButton });

        applyButton.Click += (s, e) =>
        {
            relatedGrid.EndEdit();
            foreach (DataGridViewRow gridRow in relatedGrid.Rows)
            {
                var group = (RelatedGroup)gridRow.Tag;
                group.Action = Convert.ToString(gridRow.Cells["Action"].Value);
                group.ReplaceWith = null;
                if (group.Action == "Replace")
                {
                    group.ReplaceWith = FindMeasure(Convert.ToString(gridRow.Cells["ReplaceWith"].Value));
                    if (group.ReplaceWith == null)
                    {
                        MessageBox.Show("Enter an existing measure to replace " + gridRow.Cells["Measure"].Value + " with, or choose Keep / Remove.",
                            "Replacement measure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        relatedGrid.CurrentCell = gridRow.Cells["ReplaceWith"];
                        return;
                    }
                }
            }
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };
        keepAllButton.Click += (s, e) =>
        {
            foreach (var group in relatedGroups) { group.Action = "Keep"; group.ReplaceWith = null; }
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };

        panel.Controls.Add(HeadingLabel("Related measures found in the chosen visuals. Choose what to do with each when " +
            sourceField.Label + " is replaced with " + targetField.Label + ":"), 0, 0);
        panel.Controls.Add(relatedGrid, 0, 1);
        panel.Controls.Add(hint, 0, 2);
        panel.Controls.Add(buttonRow, 0, 3);
        // No AcceptButton: Enter commits the cell being edited instead of closing.
        dlg.CancelButton = cancelButton;
        if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }
}

// ---------------------------------------------------------------------------
// Step 7: display names - one row per existing display name value
// ---------------------------------------------------------------------------
var displayTargets = new Dictionary<string, string>(StringComparer.Ordinal);   // old display name -> new ("" = remove)
Func<string, bool> IsDisplayRemoval = value =>
    string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), targetField.Name, StringComparison.Ordinal);

if (displayNameUses.Count > 0)
{
    var displayGroups = displayNameUses.GroupBy(u => u.OldDisplayName, StringComparer.Ordinal).OrderByDescending(g => g.Count()).ToList();
    using (var dlg = new Form())
    {
        var panel = CreateDialogLayout(dlg, "Set display names", 4, DialogSize(0.7, 0.6));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var displayGrid = CreateEditGrid();
        displayGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Existing", HeaderText = "Existing display name", Width = 260, ReadOnly = true });
        var targetColumn = new DataGridViewTextBoxColumn { Name = "Target", HeaderText = "Set display name (edit)", Width = 280 };
        HighlightColumn(targetColumn);
        displayGrid.Columns.Add(targetColumn);
        displayGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Result", HeaderText = "Result", Width = 260, ReadOnly = true });
        displayGrid.Columns.Add(new DataGridViewTextBoxColumn {
            Name = "Uses", HeaderText = "Uses", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 60
        });

        foreach (var g in displayGroups)
        {
            // Proposal: the existing display name with the old field name swapped for the new one
            var proposal = g.Key.IndexOf(sourceField.Name, StringComparison.OrdinalIgnoreCase) >= 0
                ? Regex.Replace(g.Key, Regex.Escape(sourceField.Name), targetField.Name.Replace("$", "$$"), RegexOptions.IgnoreCase)
                : g.Key;
            int rowIndex = displayGrid.Rows.Add(g.Key, proposal, "", g.Count() + " projection(s) in " + g.Select(u => u.Visual.File).Distinct().Count() + " visual(s)");
            displayGrid.Rows[rowIndex].Tag = g.Key;
        }

        Action<int> updateResult = rowIndex =>
        {
            var gridRow = displayGrid.Rows[rowIndex];
            var value = Convert.ToString(gridRow.Cells["Target"].Value);
            var cell = gridRow.Cells["Result"];
            if (IsDisplayRemoval(value)) { cell.Value = "Remove display name (shows " + targetField.Name + ")"; cell.Style.ForeColor = Color.DarkRed; }
            else if (string.Equals(value.Trim(), (string)gridRow.Tag, StringComparison.Ordinal)) { cell.Value = "Keep existing display name"; cell.Style.ForeColor = Color.Gray; }
            else { cell.Value = "Set display name"; cell.Style.ForeColor = Color.DarkGreen; }
        };
        for (int i = 0; i < displayGrid.Rows.Count; i++) updateResult(i);
        displayGrid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) updateResult(e.RowIndex); };
        dlg.Shown += (s, e) =>
        {
            if (displayGrid.Rows.Count == 0) return;
            displayGrid.CurrentCell = displayGrid.Rows[0].Cells["Target"];
            displayGrid.BeginEdit(true);
        };

        var hint = new Label {
            Text = "Only the highlighted column is editable. Leave it blank or equal to '" + targetField.Name + "' to REMOVE the display name so the visual shows the field name.",
            AutoSize = true, MaximumSize = new Size(1100, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0)
        };
        var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
        var applyButton = new Button { Text = "Apply", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttonRow.Controls.AddRange(new Control[] { applyButton, cancelButton });
        applyButton.Click += (s, e) =>
        {
            displayGrid.EndEdit();
            foreach (DataGridViewRow gridRow in displayGrid.Rows)
            {
                var value = Convert.ToString(gridRow.Cells["Target"].Value);
                displayTargets[(string)gridRow.Tag] = IsDisplayRemoval(value) ? "" : value.Trim();
            }
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };

        panel.Controls.Add(HeadingLabel("Set display name: " + sourceField.Label + " has explicit display names in the chosen visuals. Edit the highlighted column for " + targetField.Name + ":"), 0, 0);
        panel.Controls.Add(displayGrid, 0, 1);
        panel.Controls.Add(hint, 0, 2);
        panel.Controls.Add(buttonRow, 0, 3);
        dlg.CancelButton = cancelButton;
        if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    }
}

// ---------------------------------------------------------------------------
// Step 8: apply everything in memory and build the preview
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
preview.Columns.Add("Visual");
preview.Columns.Add("Change");
preview.Columns.Add("Old");
preview.Columns.Add("New");
preview.Columns.Add("Note");

// Sort entries on a tooltip measure that is no longer projected: point them to the
// replacement measure, or remove them when the measure was removed.
Action<VisualDoc, string, string, Measure> FixSort = (doc, table, name, replacement) =>
{
    var sortArray = doc.Root["visual"]?["query"]?["sortDefinition"]?["sort"] as JArray;
    if (sortArray == null) return;
    bool stillProjected = doc.Root["visual"]?["query"]?["queryState"]?.SelectTokens("$..Measure").Any(n => IsMeasureNode(n, table, name)) == true;
    if (stillProjected) return;
    foreach (var sortEntry in sortArray.OfType<JObject>().Where(se => IsMeasureNode(se["field"]?["Measure"], table, name)).ToList())
    {
        if (replacement == null) sortEntry.Remove();
        else
        {
            sortEntry["field"]["Measure"]["Expression"]["SourceRef"]["Entity"] = replacement.Table.Name;
            sortEntry["field"]["Measure"]["Property"] = replacement.Name;
        }
        doc.Changes++;
    }
};

// 8a. Related measures
foreach (var group in relatedGroups.Where(g => g.Action != "Keep"))
{
    var oldLabel = group.MeasureTable + "[" + group.MeasureName + "]";
    var newLabel = group.ReplaceWith != null ? group.ReplaceWith.Table.Name + "[" + group.ReplaceWith.Name + "]" : "(removed)";
    foreach (var use in group.Uses)
    {
        var doc = use.Visual;
        if (use.Usage == "Tooltip")
        {
            if (group.Action == "Remove")
            {
                if (use.Projection.Parent != null) use.Projection.Remove();
                // projections array -> JProperty "projections" -> role object -> JProperty "Tooltips"
                if (use.ProjectionArray.Count == 0) (use.ProjectionArray.Parent?.Parent?.Parent as JProperty)?.Remove();
            }
            else
            {
                var measureNode = use.Projection["field"]?["Measure"] as JObject;
                measureNode["Expression"]["SourceRef"]["Entity"] = group.ReplaceWith.Table.Name;
                measureNode["Property"] = group.ReplaceWith.Name;
                use.Projection["queryRef"] = group.ReplaceWith.Table.Name + "." + group.ReplaceWith.Name;
                if (string.Equals((string)use.Projection["nativeQueryRef"], group.MeasureName, StringComparison.Ordinal))
                    use.Projection["nativeQueryRef"] = group.ReplaceWith.Name;
            }
            FixSort(doc, group.MeasureTable, group.MeasureName, group.ReplaceWith);
        }
        else
        {
            if (group.Action == "Remove")
            {
                if (use.Properties.Property(use.PropertyName) != null) use.Properties.Remove(use.PropertyName);
                if (use.Properties.Count == 0 && use.Entry.Parent != null)
                {
                    var entryArray = (JArray)use.Entry.Parent;
                    use.Entry.Remove();
                    if (entryArray.Count == 0 && use.Container.Property(use.ObjectName) != null) use.Container.Remove(use.ObjectName);
                }
            }
            else
            {
                foreach (var measureNode in use.MeasureNodes)
                {
                    measureNode["Expression"]["SourceRef"]["Entity"] = group.ReplaceWith.Table.Name;
                    measureNode["Property"] = group.ReplaceWith.Name;
                }
            }
        }
        doc.Changes++;
        preview.Rows.Add(doc.FullLabel, group.Usage + " (" + use.Detail + ")", oldLabel, newLabel, "");
    }
}

// 8b. Display names (projections are still the same JObjects after replacement)
foreach (var use in displayNameUses)
{
    string newDisplayName;
    if (!displayTargets.TryGetValue(use.OldDisplayName, out newDisplayName)) continue;
    if (newDisplayName == "")
    {
        use.Projection.Remove("displayName");
        preview.Rows.Add(use.Visual.FullLabel, "Display name", use.OldDisplayName, "(removed - shows " + targetField.Name + ")", "");
        use.Visual.Changes++;
    }
    else if (!string.Equals(newDisplayName, use.OldDisplayName, StringComparison.Ordinal))
    {
        use.Projection["displayName"] = newDisplayName;
        preview.Rows.Add(use.Visual.FullLabel, "Display name", use.OldDisplayName, newDisplayName, "");
        use.Visual.Changes++;
    }
}

// 8c. The field itself
foreach (var doc in visualDocs)
{
    int fieldChanges = MutateSource(doc.Root, null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), false);
    doc.Changes += fieldChanges;
    preview.Rows.Add(doc.FullLabel, "Field (" + fieldChanges + " reference value(s))", sourceField.Label, targetField.Label, string.Join("; ", doc.Notes));
}

var filesToWrite = visualDocs.Where(v => v.Changes > 0).ToList();
if (filesToWrite.Count == 0)
{
    Info("The chosen options did not produce any PBIR changes. No changes made.");
    return;
}

int notesCount = visualDocs.Count(v => v.Notes.Count > 0);
if (!ShowGrid(
        "Preview field replacement",
        "Replace " + sourceField.Label + " with " + targetField.Label + " in " + filesToWrite.Count + " visual(s)." +
            (notesCount > 0 ? " " + notesCount + " visual(s) have notes - see the Note column." : "") +
            " Select Continue to confirm and write the changes.",
        preview,
        "Continue"))
{
    Info("Cancelled - no files were changed.");
    return;
}

var confirm = MessageBox.Show(
    "Replace " + sourceField.Label + " with " + targetField.Label + " in " + filesToWrite.Count + " visual.json file(s)?\n\n" +
    (pageLevelHits > 0 ? "Note: " + pageLevelHits + " page/report filter file(s) also use " + sourceField.Label + " and are NOT changed.\n" : "") +
    (reportExtensionHits > 0 ? "Note: report-level measures (reportExtensions.json) in " + reportExtensionHits + " report(s) mention " + nameToken + " and are NOT changed.\n" : "") +
    (pageLevelHits > 0 || reportExtensionHits > 0 ? "\n" : "") +
    "Only PBIR files are changed; the semantic model is not affected. PBIR edits cannot be undone in Tabular Editor. " +
    "If the reports are under source control, commit/sync first if required.\n\n" +
    "Afterwards, apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.",
    "Confirm field replacement",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no files were changed."); return; }

// ---------------------------------------------------------------------------
// Step 9: write files and summarize partial progress if an error occurs
// ---------------------------------------------------------------------------
int filesWritten = 0;
foreach (var doc in filesToWrite)
{
    try
    {
        File.WriteAllText(doc.File, doc.Root.ToString(Newtonsoft.Json.Formatting.Indented));
        filesWritten++;
    }
    catch (Exception ex)
    {
        Error("The update stopped while writing:\n" + doc.File + "\n\n" +
            filesWritten + " visual.json file(s) were already written before the error.\n\nError: " + ex.Message);
    }
}

Info("Replaced " + sourceField.Label + " with " + targetField.Label + " in " + filesWritten + " visual.json file(s)." +
    (relatedGroups.Any(g => g.Action != "Keep") ? "\nRelated measures: " + relatedGroups.Count(g => g.Action == "Replace") + " replaced, " + relatedGroups.Count(g => g.Action == "Remove") + " removed." : "") +
    (notesCount > 0 ? "\n" + notesCount + " visual(s) have notes to check (see the preview)." : "") +
    "\n\nOnly PBIR files were changed; the semantic model is not affected.\n" +
    "Next: apply external changes using Desktop Bridge, or close and re-open the PBIR file to see the changes in the reports.");
