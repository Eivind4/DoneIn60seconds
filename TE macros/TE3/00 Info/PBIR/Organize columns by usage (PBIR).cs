// ============================================================================
// 20 - Organize columns by usage (PBIR).cs
//
// The column version of macro 19: organizes columns into display folders by
// how they are used in PBIR reports and in the model, so it is easy to see
// what is in use, what is only needed by the model, and what can go. Changes
// are made in the model and can be undone with Ctrl+Z in Tabular Editor.
// PBIR files are only read.
//
// Workflow:
//   1. Pick a folder; every definition.pbir below it is found and you select
//      which reports to include.
//   2. The reports are scanned for column usage: visual wells, slicers,
//      filters (visual, page, report), sort, formatting, and field parameters
//      used in visuals.
//   3. One dialog to set it all up:
//        - Scope: all columns, or only the columns selected in TE
//        - Categories, checked from top to bottom (first match wins):
//            Used in visuals     referenced by the reports (directly or via a
//                                field parameter used in visuals)
//            Model dependency    used by the model only: measures, calculated
//                                columns / tables, calculation items, RLS,
//                                relationships, hierarchies, sort by, or a
//                                perspective
//            Unused              everything else (always last)
//          Per category: apply on/off and folder name.
//        - Hidden columns go to a "Hidden" subfolder of their category
//          (optional, on by default), e.g. Model dependency\Hidden.
//        - Optional (off by default): hide model-dependency columns and/or
//          unused columns. They then also go to the Hidden subfolder.
//        - Folder placement, as in macro 19:
//            Root + keep existing folders   Used in visuals\Customer   (default)
//            Root only                      Used in visuals  (original folder
//                                           kept in an annotation)
//            Subfolder of existing folder   Customer\Used in visuals
//        - Annotation ColumnUsage = <category> for model documentation.
//   4. Preview with a filter box, then apply.
//
// Re-running is safe: each column remembers its own folder (annotation
// ColumnUsageOriginalFolder) and the folder this macro gave it
// (ColumnUsageAppliedFolder). If the column is still in that folder, the next
// run starts from the remembered folder, so category names and placement can
// be changed. Columns moved by hand since are cleaned up by removing category
// folder names (also earlier names) from the start and end of the path.
//
// Not included: columns of field parameter and calculation group tables
// (macro 22 classifies those tables), and the hidden RowNumber column.
//
// Scope: only the model open in Tabular Editor and the PBIR reports in the
// selected folder are analysed - see macro 22 for what is not detected
// (reports elsewhere in the Service, Copilot / AI, Excel, notebooks, ...).
// A column used only there ends up as Unused here: check before deleting.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

const string AnnUsage = "ColumnUsage";
const string AnnOriginalFolder = "ColumnUsageOriginalFolder";   // the column's own folder, without category folders
const string AnnAppliedFolder = "ColumnUsageAppliedFolder";     // the folder this macro set last time
const string AnnKnownFolders = "ColumnUsageFolders";            // model-level: folder names used before

class ReportChoice
{
    public string ReportFolder;
    public string ReportName;
    public string ConnectionSummary;
    public bool Selected;
    public string Label { get { return ReportName + "  |  " + ConnectionSummary + "  |  " + ReportFolder; } }
}

class Category
{
    public string Key;          // Visual, Model, Unused
    public string Name;
    public string Folder;
    public bool Apply = true;
    public string DetectedBy;
}

class ColumnPlan
{
    public Column Column;
    public Category Category;
    public string Reason;
    public string NewFolder;
    public string BaseFolder;
    public bool Hide;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
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

// Field parameter and calculation group tables are not organized here
bool IsSupportTable(Table t)
{
    if (t is CalculationGroupTable) return true;
    try { return t.Columns.Any(c => c.GroupByColumns.Any()); } catch { return false; }
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
// Step 1: folder and reports
// ---------------------------------------------------------------------------
string rootFolder = null;
using (var fbd = new FolderBrowserDialog())
{
    fbd.Description = "Select a folder containing one or more PBIR .Report folders (definition.pbir files).";
    if (fbd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    rootFolder = fbd.SelectedPath;
}

var reportChoices = Directory.GetFiles(rootFolder, "definition.pbir", SearchOption.AllDirectories)
    .Select(fp => new ReportChoice {
        ReportFolder = Path.GetDirectoryName(fp),
        ReportName = CleanReportName(Path.GetDirectoryName(fp)),
        ConnectionSummary = ReadConnectionSummary(fp),
        Selected = true
    })
    .OrderBy(rc => rc.ReportName).ThenBy(rc => rc.ReportFolder).ToList();

if (reportChoices.Count == 0) { Info("No definition.pbir files were found below:\n" + rootFolder + "\n\nNo changes made."); return; }

using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Select PBIR reports to scan", 4, DialogSize(0.85, 0.85));
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
        if (!reportChoices.Any(rc => rc.Selected)) { MessageBox.Show("Select at least one report.", "No reports selected", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(Heading("Select the reports to scan for column usage. Include every report that uses the model - a column used only in a report you leave out ends up as 'Unused' (or 'Model dependency')."), 0, 0);
    panel.Controls.Add(searchBox, 0, 1);
    panel.Controls.Add(reportList, 0, 2);
    panel.Controls.Add(buttonRow, 0, 3);
    dlg.AcceptButton = okButton;
    dlg.CancelButton = cancelButton;
    rebuild();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}
var selectedReports = reportChoices.Where(rc => rc.Selected).ToList();

// ---------------------------------------------------------------------------
// Step 2: scan PBIR for column usage
// ---------------------------------------------------------------------------
var columnsByKey = new Dictionary<string, Column>(StringComparer.OrdinalIgnoreCase);
foreach (var cl in Model.AllColumns) columnsByKey[cl.Table.Name + "|" + cl.Name] = cl;
var visualDetails = new Dictionary<Column, HashSet<string>>();
int scannedFiles = 0, skippedFiles = 0;

void AddVisualDetail(Column cl, string detail)
{
    HashSet<string> details;
    if (!visualDetails.TryGetValue(cl, out details)) { details = new HashSet<string>(StringComparer.OrdinalIgnoreCase); visualDetails[cl] = details; }
    details.Add(detail);
}

void RecordColumn(string table, string field, List<string> path)
{
    Column cl;
    if (table == null || field == null || !columnsByKey.TryGetValue(table + "|" + field, out cl)) return;
    int projections = path.IndexOf("projections");
    int formattingRoot = Math.Max(path.IndexOf("objects"), path.IndexOf("visualContainerObjects"));
    if (projections > 0) AddVisualDetail(cl, "Visual (" + path[projections - 1] + ")");
    else if (path.Contains("sortDefinition")) AddVisualDetail(cl, "Sort");
    else if (path.Contains("filterConfig")) AddVisualDetail(cl, "Filter");
    else if (formattingRoot >= 0) AddVisualDetail(cl, "Formatting (" + (formattingRoot + 1 < path.Count ? path[formattingRoot + 1] : "") + ")");
    else AddVisualDetail(cl, "Visual");
}

void Walk(JToken token, List<string> path, Dictionary<string, string> aliases)
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
        if (parentKey == "Column")
        {
            var table = (string)obj["Expression"]?["SourceRef"]?["Entity"];
            var alias = (string)obj["Expression"]?["SourceRef"]?["Source"];
            if (table == null && alias != null) localAliases.TryGetValue(alias, out table);
            RecordColumn(table, (string)obj["Property"], path);
        }
        // Formatting per field (e.g. conditional formatting on a column): selector.metadata = "Table.Column"
        var selectorMetadata = (string)obj["selector"]?["metadata"];
        if (!string.IsNullOrWhiteSpace(selectorMetadata))
        {
            var dot = selectorMetadata.LastIndexOf('.');
            if (dot > 0 && dot < selectorMetadata.Length - 1)
                RecordColumn(selectorMetadata.Substring(0, dot), selectorMetadata.Substring(dot + 1), new List<string> { "objects", "selector" });
        }
        foreach (var property in obj.Properties())
        {
            path.Add(property.Name);
            Walk(property.Value, path, localAliases);
            path.RemoveAt(path.Count - 1);
        }
    }
    var array = token as JArray;
    if (array != null) foreach (var item in array) Walk(item, path, aliases);
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
            Walk(JObject.Parse(File.ReadAllText(file)), new List<string>(), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
            scannedFiles++;
        }
        catch (Exception ex) { skippedFiles++; Warning("Skipping unreadable PBIR JSON file:\n" + file + "\n\n" + ex.Message); }
    }
}

// Field parameters used in visuals: the columns they list are shown in visuals too
foreach (var calcTable in Model.Tables.OfType<CalculatedTable>())
{
    if (!calcTable.Columns.Any(cl => visualDetails.ContainsKey(cl))) continue;
    foreach (var cl in DependenciesOf(calcTable).OfType<Column>())
        AddVisualDetail(cl, "Field parameter (" + calcTable.Name + ")");
}

// ---------------------------------------------------------------------------
// Model dependencies (why a column without visual usage is still needed)
// ---------------------------------------------------------------------------
var perspectiveNames = Model.Perspectives.Select(p => p.Name).ToList();

string ShortList(IEnumerable<string> items)
{
    var list = items.Distinct().ToList();
    return string.Join(", ", list.Take(3)) + (list.Count > 3 ? " (+" + (list.Count - 3) + " more)" : "");
}

List<string> ModelDependencies(Column cl)
{
    var reasons = new List<string>();
    var referencedBy = cl.ReferencedBy.Cast<object>().ToList();
    var measures = referencedBy.OfType<Measure>().ToList();
    if (measures.Count > 0) reasons.Add("Measures: " + ShortList(measures.Select(m => "[" + m.Name + "]")));
    var calcColumns = referencedBy.OfType<CalculatedColumn>().ToList();
    if (calcColumns.Count > 0) reasons.Add("Calculated columns: " + ShortList(calcColumns.Select(c => c.Table.Name + "[" + c.Name + "]")));
    foreach (var group in referencedBy.Where(o => !(o is Measure) && !(o is CalculatedColumn)).GroupBy(o => o.GetType().Name))
    {
        var label = group.Key == "CalculatedTable" ? "Calculated tables / field parameters"
            : group.Key == "CalculationItem" ? "Calculation items"
            : group.Key == "TablePermission" ? "RLS (role filters)"
            : group.Key == "Function" ? "DAX UDFs"
            : group.Key;
        reasons.Add(label + ": " + group.Count());
    }
    if (Model.Relationships.Any(r => r.FromColumn == cl || r.ToColumn == cl)) reasons.Add("Relationship");
    var hierarchies = cl.Table.Hierarchies.Where(h => h.Levels.Any(l => l.Column == cl)).Select(h => h.Name).ToList();
    if (hierarchies.Count > 0) reasons.Add("Hierarchy: " + ShortList(hierarchies));
    var sortedColumns = cl.Table.Columns.Where(other => other != cl && other.SortByColumn == cl).Select(other => "[" + other.Name + "]").ToList();
    if (sortedColumns.Count > 0) reasons.Add("Sort by column for " + ShortList(sortedColumns));
    try
    {
        var perspectives = perspectiveNames.Where(p => cl.InPerspective[p]).ToList();
        if (perspectives.Count > 0) reasons.Add("Perspective: " + ShortList(perspectives));
    }
    catch { }
    return reasons;
}

// ---------------------------------------------------------------------------
// Categories and settings dialog
// ---------------------------------------------------------------------------
var categories = new List<Category> {
    new Category { Key = "Visual", Name = "Used in visuals",  Folder = "Used in visuals",  DetectedBy = "Visual wells, slicers, filters, sort, formatting, or a field parameter used in visuals" },
    new Category { Key = "Model",  Name = "Model dependency", Folder = "Model dependency", DetectedBy = "Measures, calculated columns/tables, calculation items, RLS, relationships, hierarchies, sort by, perspectives" },
    new Category { Key = "Unused", Name = "Unused",           Folder = "Unused",           DetectedBy = "Everything else (always last)" }
};

var candidateColumns = Model.AllColumns.Where(cl => !IsSupportTable(cl.Table) && cl.Type != ColumnType.RowNumber).ToList();
var selectedColumns = Selected.Columns.Where(cl => candidateColumns.Contains(cl)).ToList();
var classification = candidateColumns.ToDictionary(cl => cl, cl =>
{
    HashSet<string> details;
    if (visualDetails.TryGetValue(cl, out details)) return new { Key = "Visual", Reason = string.Join("; ", details.Take(4)) + (details.Count > 4 ? "; ..." : "") };
    var deps = ModelDependencies(cl);
    if (deps.Count > 0) return new { Key = "Model", Reason = string.Join("; ", deps) };
    return new { Key = "Unused", Reason = "No usage found in the selected reports and no model dependency" };
});

bool selectedOnly = selectedColumns.Count > 0;
int placementMode = 0;                  // 0 = root + keep existing, 1 = root only, 2 = subfolder of existing
bool useHiddenFolder = true;
string hiddenFolderName = "Hidden";
bool hideModelColumns = false;
bool hideUnusedColumns = false;
bool writeAnnotation = true;

using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Organize columns by usage", 7, DialogSize(0.75, 0.8));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var scopeCheck = new CheckBox {
        Text = "Only the " + selectedColumns.Count + " column(s) selected in Tabular Editor (otherwise all " + candidateColumns.Count + " columns)",
        Checked = selectedOnly, Enabled = selectedColumns.Count > 0, AutoSize = true, Margin = new Padding(0, 0, 0, 8)
    };

    var grid = new DataGridView {
        Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect, MultiSelect = false,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false, EditMode = DataGridViewEditMode.EditOnEnter
    };
    grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
    grid.DefaultCellStyle.ForeColor = Color.FromArgb(90, 90, 90);
    var applyColumn = new DataGridViewCheckBoxColumn { Name = "Apply", HeaderText = "Apply", Width = 55 };
    var folderColumn = new DataGridViewTextBoxColumn { Name = "Folder", HeaderText = "Folder name (edit)", Width = 200 };
    Highlight(applyColumn);
    Highlight(folderColumn);
    grid.Columns.Add(applyColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "Category", Width = 150, ReadOnly = true });
    grid.Columns.Add(folderColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Count", HeaderText = "Columns", Width = 70, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Hidden", HeaderText = "Of which hidden", Width = 110, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Detected", HeaderText = "Detected by (in this order, first match wins)", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200 });
    grid.CurrentCellDirtyStateChanged += (s, e) => { if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };

    var placementBox = new GroupBox { Text = "Folder placement", AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 0) };
    var placementFlow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Dock = DockStyle.Fill };
    var modeRootKeep = new RadioButton { Text = "Category as root folder, keep existing folders below it  (Used in visuals\\Customer)", AutoSize = true, Checked = true };
    var modeRootOnly = new RadioButton { Text = "Category as root folder only  (Used in visuals) - the original folder is kept in an annotation", AutoSize = true };
    var modeSuffix = new RadioButton { Text = "Category as subfolder of the existing folder  (Customer\\Used in visuals)", AutoSize = true };
    placementFlow.Controls.AddRange(new Control[] { modeRootKeep, modeRootOnly, modeSuffix });
    placementBox.Controls.Add(placementFlow);

    var optionsFlow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, Margin = new Padding(0, 8, 0, 0) };
    var hiddenCheck = new CheckBox { Text = "Hidden columns in a subfolder of their category:", AutoSize = true, Checked = true, Margin = new Padding(0, 4, 4, 0) };
    var hiddenBox = new TextBox { Text = hiddenFolderName, Width = 140, BackColor = editBack };
    var annotationCheck = new CheckBox { Text = "Write annotation " + AnnUsage + " = <category> (for model documentation)", AutoSize = true, Checked = true, Margin = new Padding(24, 4, 0, 0) };
    var hideModelCheck = new CheckBox { Text = "Hide Model dependency columns (not used in visuals)", AutoSize = true, Checked = false, Margin = new Padding(0, 8, 24, 0) };
    var hideUnusedCheck = new CheckBox { Text = "Hide Unused columns", AutoSize = true, Checked = false, Margin = new Padding(0, 8, 0, 0) };
    optionsFlow.Controls.AddRange(new Control[] { hiddenCheck, hiddenBox, annotationCheck, hideModelCheck, hideUnusedCheck });
    optionsFlow.SetFlowBreak(annotationCheck, true);

    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
    var previewButton = new Button { Text = "Preview", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { previewButton, cancelButton });

    bool loading = false;
    Action refreshCounts = () =>
    {
        var inScope = scopeCheck.Checked ? selectedColumns : candidateColumns;
        foreach (DataGridViewRow row in grid.Rows)
        {
            var cat = (Category)row.Tag;
            var inCategory = inScope.Where(cl => classification[cl].Key == cat.Key).ToList();
            row.Cells["Count"].Value = inCategory.Count;
            row.Cells["Hidden"].Value = inCategory.Count(cl => cl.IsHidden);
        }
    };
    loading = true;
    foreach (var cat in categories)
    {
        int index = grid.Rows.Add(cat.Apply, cat.Name, cat.Folder, "", "", cat.DetectedBy);
        grid.Rows[index].Tag = cat;
    }
    loading = false;
    grid.CellValueChanged += (s, e) =>
    {
        if (loading || e.RowIndex < 0) return;
        var row = grid.Rows[e.RowIndex];
        var cat = (Category)row.Tag;
        cat.Apply = Convert.ToBoolean(row.Cells["Apply"].Value);
        cat.Folder = Convert.ToString(row.Cells["Folder"].Value).Trim().Replace("/", "\\").Trim('\\');
    };
    scopeCheck.CheckedChanged += (s, e) => refreshCounts();
    hiddenCheck.CheckedChanged += (s, e) => hiddenBox.Enabled = hiddenCheck.Checked;

    previewButton.Click += (s, e) =>
    {
        grid.EndEdit();
        foreach (var cat in categories.Where(c => c.Apply))
            if (string.IsNullOrWhiteSpace(cat.Folder)) { MessageBox.Show("Enter a folder name for '" + cat.Name + "', or untick Apply.", "Folder name", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        if (hiddenCheck.Checked && string.IsNullOrWhiteSpace(hiddenBox.Text)) { MessageBox.Show("Enter a subfolder name for hidden columns, or untick the option.", "Hidden subfolder", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        selectedOnly = scopeCheck.Checked;
        placementMode = modeRootKeep.Checked ? 0 : modeRootOnly.Checked ? 1 : 2;
        useHiddenFolder = hiddenCheck.Checked;
        hiddenFolderName = hiddenBox.Text.Trim().Replace("/", "\\").Trim('\\');
        hideModelColumns = hideModelCheck.Checked;
        hideUnusedColumns = hideUnusedCheck.Checked;
        writeAnnotation = annotationCheck.Checked;
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(Heading("Organize columns into folders by how they are used in the " + selectedReports.Count + " selected report(s) (" + scannedFiles + " PBIR files scanned" + (skippedFiles > 0 ? ", " + skippedFiles + " skipped" : "") + ") and in the model."), 0, 0);
    panel.Controls.Add(scopeCheck, 0, 1);
    panel.Controls.Add(grid, 0, 2);
    panel.Controls.Add(placementBox, 0, 3);
    panel.Controls.Add(optionsFlow, 0, 4);
    panel.Controls.Add(new Label {
        Text = "Untick Apply to leave that category's columns in their current folder. Field parameter and calculation group tables are not included. " +
               "Only the reports in the selected folder are analysed - usage elsewhere (Service, Copilot, Excel, notebooks) is not detected, so check before deleting Unused columns.",
        AutoSize = true, MaximumSize = new Size(1200, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 0)
    }, 0, 5);
    panel.Controls.Add(buttonRow, 0, 6);
    dlg.CancelButton = cancelButton;
    refreshCounts();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 3: compute new folders
// ---------------------------------------------------------------------------
// Folder names to strip from columns moved by hand: current and earlier category folders
var knownFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
foreach (var cat in categories) if (!string.IsNullOrWhiteSpace(cat.Folder)) foreach (var part in cat.Folder.Split('\\')) knownFolders.Add(part.Trim());
if (useHiddenFolder && !string.IsNullOrWhiteSpace(hiddenFolderName)) foreach (var part in hiddenFolderName.Split('\\')) knownFolders.Add(part.Trim());
var previousFolders = Model.GetAnnotation(AnnKnownFolders);
if (!string.IsNullOrWhiteSpace(previousFolders)) foreach (var part in previousFolders.Split('|')) if (part.Trim().Length > 0) knownFolders.Add(part.Trim());

// The column's own folder, without category folders (same rules as macro 19)
string BaseFolder(Column cl)
{
    var current = cl.DisplayFolder ?? "";
    var applied = cl.GetAnnotation(AnnAppliedFolder);
    if (applied != null && string.Equals(applied, current, StringComparison.OrdinalIgnoreCase))
        return cl.GetAnnotation(AnnOriginalFolder) ?? "";

    var parts = current.Split('\\').Select(p => p.Trim()).Where(p => p.Length > 0).ToList();
    while (parts.Count > 0 && knownFolders.Contains(parts[0])) parts.RemoveAt(0);
    while (parts.Count > 0 && knownFolders.Contains(parts[parts.Count - 1])) parts.RemoveAt(parts.Count - 1);
    var basePath = string.Join("\\", parts);
    if (basePath.Length == 0) basePath = cl.GetAnnotation(AnnOriginalFolder) ?? "";
    return basePath;
}

string Combine(params string[] parts) { return string.Join("\\", parts.Where(p => !string.IsNullOrWhiteSpace(p))); }

var inScopeColumns = selectedOnly ? selectedColumns : candidateColumns;
var plans = new List<ColumnPlan>();
foreach (var cl in inScopeColumns)
{
    var info = classification[cl];
    var cat = categories.First(c => c.Key == info.Key);
    var plan = new ColumnPlan { Column = cl, Category = cat, Reason = info.Reason };
    plan.Hide = !cl.IsHidden && ((cat.Key == "Model" && hideModelColumns) || (cat.Key == "Unused" && hideUnusedColumns));
    bool hiddenAfter = cl.IsHidden || plan.Hide;

    if (!cat.Apply) plan.NewFolder = cl.DisplayFolder ?? "";
    else
    {
        var categoryFolder = Combine(cat.Folder, useHiddenFolder && hiddenAfter ? hiddenFolderName : null);
        var basePath = BaseFolder(cl);
        plan.BaseFolder = basePath;
        plan.NewFolder = placementMode == 0 ? Combine(categoryFolder, basePath)
            : placementMode == 1 ? categoryFolder
            : Combine(basePath, categoryFolder);
    }
    plans.Add(plan);
}

// ---------------------------------------------------------------------------
// Step 4: preview with filter box
// ---------------------------------------------------------------------------
var preview = new System.Data.DataTable();
foreach (var column in new[] { "Table", "Column", "Category", "Hidden", "Current folder", "New folder", "Changed", "Why" }) preview.Columns.Add(column);
foreach (var plan in plans.OrderBy(p => categories.IndexOf(p.Category)).ThenBy(p => p.Column.Table.Name).ThenBy(p => p.Column.Name))
    preview.Rows.Add(plan.Column.Table.Name, plan.Column.Name,
        plan.Category.Name + (plan.Category.Apply ? "" : " (not applied)"),
        plan.Hide ? "Will be hidden" : plan.Column.IsHidden ? "Yes" : "No",
        plan.Column.DisplayFolder ?? "", plan.NewFolder,
        string.Equals(plan.NewFolder, plan.Column.DisplayFolder ?? "", StringComparison.Ordinal) && !plan.Hide ? "No" : "Yes",
        plan.Reason ?? "");

int changes = plans.Count(p => !string.Equals(p.NewFolder, p.Column.DisplayFolder ?? "", StringComparison.Ordinal));
int hides = plans.Count(p => p.Hide);
using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Preview - organize columns", 4, DialogSize(0.9, 0.85));
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
            parts.Add("(" + string.Join(" OR ", new[] { "[Table]", "[Column]", "[Category]", "[Hidden]", "[Current folder]", "[New folder]", "[Why]" }
                .Select(c => c + " LIKE '%" + text + "%'")) + ")");
        if (changedOnly.Checked) parts.Add("Changed = 'Yes'");
        view.RowFilter = string.Join(" AND ", parts);
    };
    filterBox.TextChanged += (s, e) => applyFilter();
    changedOnly.CheckedChanged += (s, e) => applyFilter();

    var buttonRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var applyButton = new Button { Text = "Apply", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0), Enabled = changes > 0 || hides > 0 || writeAnnotation };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttonRow.Controls.AddRange(new Control[] { applyButton, cancelButton });

    var summaryText = string.Join("   ", plans.GroupBy(p => p.Category).OrderBy(g => categories.IndexOf(g.Key))
        .Select(g => g.Key.Name + ": " + g.Count() + " (" + g.Count(p => p.Column.IsHidden || p.Hide) + " hidden)"));
    panel.Controls.Add(Heading(plans.Count + " column(s), " + changes + " move to another folder" + (hides > 0 ? ", " + hides + " will be hidden" : "") + ". Click a column header to sort.\n" + summaryText), 0, 0);
    panel.Controls.Add(filterRow, 0, 1);
    panel.Controls.Add(previewGrid, 0, 2);
    panel.Controls.Add(buttonRow, 0, 3);
    dlg.CancelButton = cancelButton;
    dlg.Shown += (s, e) => filterBox.Focus();
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 5: apply (undoable with Ctrl+Z in Tabular Editor)
// ---------------------------------------------------------------------------
int moved = 0, hidden = 0;
foreach (var plan in plans)
{
    var cl = plan.Column;
    if (plan.Hide) { cl.IsHidden = true; hidden++; }
    if (plan.Category.Apply)
    {
        // Remember the column's own folder and the folder set now, so a later run
        // (also with other category names or placement) starts from the right folder
        if (plan.BaseFolder.Length > 0) cl.SetAnnotation(AnnOriginalFolder, plan.BaseFolder);
        else cl.RemoveAnnotation(AnnOriginalFolder);
        cl.SetAnnotation(AnnAppliedFolder, plan.NewFolder);

        if (!string.Equals(plan.NewFolder, cl.DisplayFolder ?? "", StringComparison.Ordinal))
        {
            cl.DisplayFolder = plan.NewFolder;
            moved++;
        }
    }
    if (writeAnnotation) cl.SetAnnotation(AnnUsage, plan.Category.Name);
}

// Remember the folder names so a later run with other names can still clean them up
Model.SetAnnotation(AnnKnownFolders, string.Join("|", knownFolders.OrderBy(x => x)));

Info("Done. Moved " + moved + " column(s) to category folders" + (hidden > 0 ? ", hid " + hidden + " column(s)" : "") +
    (writeAnnotation ? " and wrote the " + AnnUsage + " annotation on " + plans.Count + " column(s)" : "") + ".\n\n" +
    "Undo with Ctrl+Z in Tabular Editor. Save the model to keep the changes.\n" +
    "Tip: in model documentation, filter on the " + AnnUsage + " annotation to show only columns in use.");
