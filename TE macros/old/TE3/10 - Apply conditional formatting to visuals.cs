// ============================================================================
// 10 - Apply conditional formatting to visuals.cs
//
// Tabular Editor 3 C# macro that applies "Field value" conditional
// formatting to PBIR visual.json files, driven by ONE color measure per run
// (e.g. "CF text Sales Amount", created by 5 - Create Conditional formatting
// DaxUDF.cs + 5 - Create measure from model dependent functions.cs).
//
// TE3-only: dialogs use Dock/TableLayoutPanel-based dynamic layout and are
// sized to ~85% of the screen by default (no fixed pixel positions from the
// TE2 days), so nothing gets clipped and every form is resizable and
// maximizable.
//
// Scope (kept intentionally simple for now):
//   - Table / Matrix (tableEx, pivotTable): Background or Font color of a
//     field in the Values well. Category-value scoped formatting is not
//     included in this macro.
//   - Bar / Column charts (all clustered/stacked/100% variants, in both bar
//     and column orientation): Background (bar/column fill), only offered
//     when the visual has EXACTLY ONE measure/field in its value role (Y),
//     since that's the field the fill color rule targets unambiguously.
//   - Advanced card (cardVisual / advancedCard): field-scoped value, label,
//     reference label value/title/detail colors, plus default card fill.
//
// Workflow:
//   Step 1: pick the report (its definition.pbir file, or the .Report
//           folder directly).
//   Step 2: pick the color measure.
//             - If exactly one measure is selected in the model tree when
//               the macro is run, that measure is used automatically.
//             - Otherwise (none or multiple selected), a searchable list
//               box is shown to pick exactly one.
//   Step 3: a MULTI-SELECT (checkbox) tree lets you tick as many targets as
//           you like in one pass:
//             Page > Visual type > Visual > Element > Field
//           Element is simply "Background" or "Font" - no other labels.
//           The dialog has an overall search box plus dedicated filters for
//           Visual type, Title, and Element (Background / Font), so large
//           reports stay easy to navigate.
//   Step 4: preview every change in one table, confirm once, then write
//           all of them into the relevant visual.json files. Safe to
//           re-run - an existing entry for the same field/scope is updated
//           in place rather than duplicated.
//
// IMPORTANT (per Tabular Editor's PBIR scripting guidance):
//   - Close Power BI Desktop before running this script.
//   - Commit your current changes to git first - PBIR edits made by a
//     script are NOT undoable from within Tabular Editor.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

ScriptHelper.WaitFormVisible = false;

// ---------------------------------------------------------------------------
// Data carried from tree leaf -> write step
// ---------------------------------------------------------------------------
public class LeafConfig
{
    public string Key;             // stable identity used to persist checkbox state across filter/rebuild
    public string File;
    public JObject Root;
    public string ContainerName;   // "values" | "dataPoint"
    public string PropertyKey;     // "backColor" | "fontColor" | "fill"
    public string FieldEntity;
    public string FieldProperty;
    public int? MatchingOption;
    public bool DefaultSelector;
    public JObject SelectorTemplate;
    public string PageDisplay;
    public string VisualType;
    public string VisualLabel;
    public string ElementLabel;    // "Background" | "Font"
    public string DisplayPath;
}

// ---------------------------------------------------------------------------
// UI helpers - fully dynamic (Dock/TableLayoutPanel), TE3-only, resizable,
// sized generously by default (no more cramped TE2-era dialogs).
// ---------------------------------------------------------------------------
Func<Size> bigDialogSize = () =>
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * 0.85), (int)(wa.Height * 0.85));
};

bool cancelled = false;

// Searchable single-select list. Returns the Key of the chosen item.
Func<List<(string Label, string Key)>, string, string> pickOne = (items, title) =>
{
    string chosen = null;
    using (var form = new Form())
    {
        form.Text = title;
        form.Size = bigDialogSize();
        form.MinimumSize = new Size(560, 420);
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.Sizable;
        form.MaximizeBox = true;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var searchBox = new TextBox { Dock = DockStyle.Fill, Font = font, Margin = new Padding(0, 0, 0, 8) };
        var lb = new ListBox { Dock = DockStyle.Fill, Font = font, HorizontalScrollbar = true, IntegralHeight = false };

        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
        var okButton = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel, Font = font };
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);

        Action<string> rebuild = (filter) =>
        {
            lb.Items.Clear();
            var f = (filter ?? "").Trim();
            foreach (var it in items)
                if (string.IsNullOrEmpty(f) || it.Label.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                    lb.Items.Add(it.Label);
            if (lb.Items.Count > 0) lb.SelectedIndex = 0;
        };
        searchBox.TextChanged += (s, e) => rebuild(searchBox.Text);
        rebuild(null);

        layout.Controls.Add(searchBox, 0, 0);
        layout.Controls.Add(lb, 0, 1);
        layout.Controls.Add(buttonPanel, 0, 2);
        form.Controls.Add(layout);
        form.AcceptButton = okButton;
        form.CancelButton = cancelButton;

        if (form.ShowDialog() != DialogResult.OK) { cancelled = true; return null; }
        if (lb.SelectedIndex < 0) { cancelled = true; return null; }
        var label = (string)lb.Items[lb.SelectedIndex];
        chosen = items.First(it => it.Label == label).Key;
    }
    return chosen;
};

// Multi-select (checkbox) tree with an overall search box plus dedicated
// Visual type / Title / Element filters. Returns the List<LeafConfig> that
// ended up checked, regardless of which filter was active when checked.
Func<List<LeafConfig>, string, List<LeafConfig>> pickManyLeaves = (allLeaves, title) =>
{
    var checkedKeys = new HashSet<string>();
    var result = new List<LeafConfig>();

    using (var form = new Form())
    {
        form.Text = title;
        form.Size = bigDialogSize();
        form.MinimumSize = new Size(720, 520);
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.Sizable;
        form.MaximizeBox = true;
        var font = new Font("Segoe UI", 9);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // search
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // filters
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // tree
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // status
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // buttons

        // --- Row 0: overall search box ---
        var searchPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        var searchLabel = new Label { Text = "Search:", AutoSize = true, Font = font, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) };
        var searchBox = new TextBox { Dock = DockStyle.Fill, Font = font };
        searchPanel.Controls.Add(searchLabel, 0, 0);
        searchPanel.Controls.Add(searchBox, 1, 0);

        // --- Row 1: Visual type / Title / Element filters ---
        var filterPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Height = 160, Margin = new Padding(0, 0, 0, 8) };
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));

        var visualTypeGroup = new GroupBox { Text = "Visual type", Dock = DockStyle.Fill, Font = font };
        var visualTypeList = new CheckedListBox { Dock = DockStyle.Fill, Font = font, CheckOnClick = true };
        foreach (var vt in allLeaves.Select(l => l.VisualType).Distinct().OrderBy(x => x))
            visualTypeList.Items.Add(vt, true);
        visualTypeGroup.Controls.Add(visualTypeList);

        var elementGroup = new GroupBox { Text = "Element", Dock = DockStyle.Fill, Font = font };
        var elementList = new CheckedListBox { Dock = DockStyle.Fill, Font = font, CheckOnClick = true };
        foreach (var el in allLeaves.Select(l => l.ElementLabel).Distinct().OrderBy(x => x))
            elementList.Items.Add(el, true);
        elementGroup.Controls.Add(elementList);

        var titleGroup = new GroupBox { Text = "Title contains", Dock = DockStyle.Fill, Font = font };
        var titleBox = new TextBox { Dock = DockStyle.Top, Font = font };
        titleGroup.Controls.Add(titleBox);

        filterPanel.Controls.Add(visualTypeGroup, 0, 0);
        filterPanel.Controls.Add(elementGroup, 1, 0);
        filterPanel.Controls.Add(titleGroup, 2, 0);

        // --- Row 2: the checkbox tree ---
        var tree = new TreeView { Dock = DockStyle.Fill, Font = font, Scrollable = true, CheckBoxes = true };

        // --- Row 3: status label ---
        var statusLabel = new Label { Text = "0 selected", Dock = DockStyle.Fill, AutoSize = true, Font = font, ForeColor = Color.DimGray, Margin = new Padding(0, 4, 0, 0) };

        // --- Row 4: buttons ---
        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
        var selectVisibleButton = new Button { Text = "Check all visible", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var clearAllButton = new Button { Text = "Clear all", AutoSize = true, Font = font, Margin = new Padding(0, 0, 24, 0) };
        var okButton = new Button { Text = "Apply", AutoSize = true, Font = font, Margin = new Padding(0, 0, 8, 0) };
        var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel, Font = font };
        buttonPanel.Controls.Add(selectVisibleButton);
        buttonPanel.Controls.Add(clearAllButton);
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);

        Action updateStatus = () => statusLabel.Text = checkedKeys.Count + " selected";

        // Recursively propagate a check-state change down to all descendant
        // leaf nodes, keeping checkedKeys as the single source of truth.
        bool suppress = false;
        Action<TreeNode, bool> setCheckedRecursive = null;
        setCheckedRecursive = (node, isChecked) =>
        {
            node.Checked = isChecked;
            if (node.Tag is LeafConfig lc)
            {
                if (isChecked) checkedKeys.Add(lc.Key); else checkedKeys.Remove(lc.Key);
            }
            foreach (TreeNode child in node.Nodes) setCheckedRecursive(child, isChecked);
        };
        tree.AfterCheck += (s, e) =>
        {
            if (suppress) return;
            suppress = true;
            setCheckedRecursive(e.Node, e.Node.Checked);
            suppress = false;
            updateStatus();
        };

        Func<List<LeafConfig>> currentFiltered = () =>
        {
            var visualTypes = new HashSet<string>(visualTypeList.CheckedItems.Cast<string>());
            var elements = new HashSet<string>(elementList.CheckedItems.Cast<string>());
            var search = (searchBox.Text ?? "").Trim();
            var titleFilter = (titleBox.Text ?? "").Trim();
            return allLeaves.Where(l =>
                visualTypes.Contains(l.VisualType) &&
                elements.Contains(l.ElementLabel) &&
                (string.IsNullOrEmpty(titleFilter) || l.VisualLabel.IndexOf(titleFilter, StringComparison.OrdinalIgnoreCase) >= 0) &&
                (string.IsNullOrEmpty(search) || l.DisplayPath.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();
        };

        Action rebuildTree = () =>
        {
            tree.Nodes.Clear();
            var filtered = currentFiltered();

            foreach (var pageGroup in filtered.GroupBy(l => l.PageDisplay))
            {
                var pageNode = new TreeNode(pageGroup.Key);
                foreach (var typeGroup in pageGroup.GroupBy(l => l.VisualType))
                {
                    var typeNode = new TreeNode(typeGroup.Key);
                    foreach (var visualGroup in typeGroup.GroupBy(l => l.VisualLabel))
                    {
                        var visualNode = new TreeNode(visualGroup.Key);
                        foreach (var elGroup in visualGroup.GroupBy(l => l.ElementLabel))
                        {
                            var elNode = new TreeNode(elGroup.Key);

                            foreach (var l in elGroup)
                            {
                                var n = new TreeNode(l.DefaultSelector ? "Visual" : l.FieldEntity + "." + l.FieldProperty) { Tag = l, Checked = checkedKeys.Contains(l.Key) };
                                elNode.Nodes.Add(n);
                            }

                            if (elNode.Nodes.Count > 0) visualNode.Nodes.Add(elNode);
                        }
                        if (visualNode.Nodes.Count > 0) typeNode.Nodes.Add(visualNode);
                    }
                    if (typeNode.Nodes.Count > 0) pageNode.Nodes.Add(typeNode);
                }
                if (pageNode.Nodes.Count > 0) tree.Nodes.Add(pageNode);
            }
            bool hasTextFilter = !string.IsNullOrWhiteSpace(searchBox.Text) || !string.IsNullOrWhiteSpace(titleBox.Text);
            if (hasTextFilter) tree.ExpandAll(); else tree.CollapseAll();
            updateStatus();
        };

        searchBox.TextChanged += (s, e) => rebuildTree();
        titleBox.TextChanged += (s, e) => rebuildTree();
        visualTypeList.ItemCheck += (s, e) => form.BeginInvoke((Action)(() => rebuildTree()));
        elementList.ItemCheck += (s, e) => form.BeginInvoke((Action)(() => rebuildTree()));

        selectVisibleButton.Click += (s, e) =>
        {
            foreach (var l in currentFiltered()) checkedKeys.Add(l.Key);
            rebuildTree();
        };
        clearAllButton.Click += (s, e) =>
        {
            checkedKeys.Clear();
            rebuildTree();
        };
        okButton.Click += (s, e) =>
        {
            if (checkedKeys.Count == 0)
            {
                MessageBox.Show("Please check at least one target.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        layout.Controls.Add(searchPanel, 0, 0);
        layout.Controls.Add(filterPanel, 0, 1);
        layout.Controls.Add(tree, 0, 2);
        layout.Controls.Add(statusLabel, 0, 3);
        layout.Controls.Add(buttonPanel, 0, 4);
        form.Controls.Add(layout);
        form.CancelButton = cancelButton;

        rebuildTree();

        if (form.ShowDialog() != DialogResult.OK) { cancelled = true; return null; }
        result = allLeaves.Where(l => checkedKeys.Contains(l.Key)).ToList();
    }
    return result;
};

// ---------------------------------------------------------------------------
// Step 1: locate the report (definition.pbir) - FIRST
// ---------------------------------------------------------------------------
string reportFolder = null;
using (var ofd = new OpenFileDialog())
{
    ofd.Title = "Select the report's definition.pbir file";
    ofd.Filter = "Report definition (definition.pbir)|definition.pbir|All files (*.*)|*.*";
    if (ofd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no report selected."); return; }
    reportFolder = Path.GetDirectoryName(ofd.FileName);
}

string pagesFolder = Path.Combine(reportFolder, "definition", "pages");
if (!Directory.Exists(pagesFolder)) Error("definition\\pages not found under: " + reportFolder);

// ---------------------------------------------------------------------------
// Visual-type classification - table/matrix, bar/column charts, and advanced cards only
// ---------------------------------------------------------------------------
var tableMatrixTypes = new HashSet<string> { "tableEx", "pivotTable" };
var barColumnTypes = new HashSet<string>
{
    "barChart", "columnChart", "clusteredBarChart", "clusteredColumnChart",
    "stackedBarChart", "stackedColumnChart",
    "hundredPercentStackedBarChart", "hundredPercentStackedColumnChart"
};
var advancedCardTypes = new HashSet<string> { "cardVisual", "advancedCard" };

// Roles that behave as "value" fields (matrix/table Values, chart Y).
var valueRoleNames = new HashSet<string> { "Values", "Y", "Y2", "Data" };

Func<JToken, List<(string Entity, string Property, string Role, bool IsMeasure)>> extractRoleFields = (queryState) =>
{
    var found = new List<(string, string, string, bool)>();
    if (!(queryState is JObject qsObj)) return found;
    foreach (var roleProp in qsObj.Properties())
    {
        var projections = roleProp.Value?["projections"] as JArray;
        if (projections == null) continue;
        foreach (var proj in projections)
        {
            var field = proj["field"];
            var isMeasure = field?["Measure"] is JObject;
            var measure = field?["Measure"] as JObject ?? field?["Column"] as JObject;
            if (measure == null) continue;
            var sourceRef = measure["Expression"]?["SourceRef"] as JObject;
            if (sourceRef == null || sourceRef["Schema"] != null) continue; // skip extension measures
            var entity = (string)sourceRef["Entity"];
            var property = (string)measure["Property"];
            if (entity != null && property != null) found.Add((entity, property, roleProp.Name, isMeasure));
        }
    }
    return found.Distinct().ToList();
};

// ---------------------------------------------------------------------------
// Scan every visual.json under the report
// ---------------------------------------------------------------------------
var pageFolders = Directory.GetDirectories(pagesFolder)
    .Where(d => File.Exists(Path.Combine(d, "page.json")))
    .OrderBy(d => d)
    .ToList();
if (pageFolders.Count == 0) Error("No pages found under: " + pagesFolder);

var allVisuals = new List<(string File, JObject Root, string PageDisplay, string VisualType, string Label,
    List<(string Entity, string Property)> ValueFields)>();

foreach (var pageFolder in pageFolders)
{
    var visualsFolder = Path.Combine(pageFolder, "visuals");
    if (!Directory.Exists(visualsFolder)) continue;

    JObject pageJson = null;
    try { pageJson = JObject.Parse(File.ReadAllText(Path.Combine(pageFolder, "page.json"))); } catch { }
    var pageDisplayName = (string)pageJson?["displayName"] ?? Path.GetFileName(pageFolder);

    foreach (var file in Directory.GetFiles(visualsFolder, "visual.json", SearchOption.AllDirectories))
    {
        JObject root;
        try { root = JObject.Parse(File.ReadAllText(file)); } catch { continue; }

        var visualType = (string)root["visual"]?["visualType"];
        if (string.IsNullOrEmpty(visualType)) continue;
        if (!tableMatrixTypes.Contains(visualType) && !barColumnTypes.Contains(visualType) && !advancedCardTypes.Contains(visualType)) continue; // out of scope for now

        var objectId = (string)root["name"] ?? Path.GetFileName(Path.GetDirectoryName(file));
        var queryState = root["visual"]?["query"]?["queryState"];
        var roleFields = extractRoleFields(queryState);

        var valueRoleFields = roleFields.Where(f => valueRoleNames.Contains(f.Role)).ToList();
        var valueFields = tableMatrixTypes.Contains(visualType)
            ? valueRoleFields.Select(f => (f.Entity, f.Property)).Distinct().ToList()
            : valueRoleFields.Where(f => f.IsMeasure).Select(f => (f.Entity, f.Property)).Distinct().ToList();
        // Bar/column charts only qualify when exactly one measure/field
        // drives the value axis (X for bar, Y for column) - otherwise the
        // fill-color rule would be ambiguous about which field it colors.
        if (barColumnTypes.Contains(visualType) && valueFields.Count != 1) continue;
        if (valueFields.Count == 0) continue;

        Func<string, string> literalText = (raw) => raw == null ? null : raw.Trim('\'');
        var title = literalText((string)root["visual"]?["visualContainerObjects"]?["title"]?[0]?["properties"]?["text"]?["expr"]?["Literal"]?["Value"]);
        var subtitle = literalText((string)root["visual"]?["visualContainerObjects"]?["title"]?[0]?["properties"]?["subTitle"]?["expr"]?["Literal"]?["Value"]);
        if (subtitle == null)
            subtitle = literalText((string)root["visual"]?["visualContainerObjects"]?["subTitle"]?[0]?["properties"]?["text"]?["expr"]?["Literal"]?["Value"]);

        string label = !string.IsNullOrWhiteSpace(title)
            ? title + (string.IsNullOrWhiteSpace(subtitle) ? "" : " / " + subtitle)
            : objectId + "  -> " + string.Join(", ", valueFields.Select(f => f.Entity + "." + f.Property).Take(3));

        allVisuals.Add((file, root, pageDisplayName, visualType, label, valueFields));
    }
}
if (allVisuals.Count == 0) Error("No eligible table/matrix, bar/column, or advanced card visuals found under: " + pagesFolder);

// ---------------------------------------------------------------------------
// Step 2: pick the color measure - use the single selected measure if
// there is exactly one, otherwise fall back to a searchable pick list.
// ---------------------------------------------------------------------------
Measure colorMeasure = null;
if (Selected.Measures.Count() == 1)
{
    colorMeasure = Selected.Measures.First();
    Info("Using selected measure: " + colorMeasure.DaxObjectFullName);
}
else
{
    var measureItems = Model.AllMeasures
        .OrderBy(m => m.Table.Name).ThenBy(m => m.Name)
        .Select(m => (m.Table.Name + " | " + m.Name, m.Table.Name + "." + m.Name))
        .ToList();
    string colorMeasureKey = pickOne(measureItems, "Pick the color measure to apply as conditional formatting");
    if (cancelled || colorMeasureKey == null) { Info("Cancelled - no files were changed."); return; }
    var colorParts = colorMeasureKey.Split(new[] { '.' }, 2);
    colorMeasure = Model.AllMeasures.FirstOrDefault(m => m.Table.Name == colorParts[0] && m.Name == colorParts[1]);
}
if (colorMeasure == null) Error("Could not resolve a color measure - aborting.");

// Validate the selected formatting measure as well as we can before writing PBIR.
// FormatString is not reliable for this purpose; conditional formatting expects
// the measure result itself to be a text HEX color such as #1F84BB.
Func<string, bool> isHexColor = value =>
{
    if (string.IsNullOrWhiteSpace(value)) return false;
    var v = value.Trim().Trim('\'', '"');
    return Regex.IsMatch(v, "^#(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$");
};

try
{
    var validationDax = "EVALUATE ROW(\"Color\", " + colorMeasure.DaxObjectFullName + ")";
    var validationResult = ExecuteDax(validationDax);
    string sampleColor = null;
    if (validationResult != null && validationResult.Tables.Count > 0 && validationResult.Tables[0].Rows.Count > 0)
    {
        var raw = validationResult.Tables[0].Rows[0][0];
        if (raw != null && raw != DBNull.Value) sampleColor = Convert.ToString(raw);
    }

    if (!isHexColor(sampleColor))
    {
        var response = MessageBox.Show(
            "The selected measure did not evaluate to a HEX color in the current unfiltered context.\n\n"
            + "Measure: " + colorMeasure.DaxObjectFullName + "\n"
            + "Returned: " + (string.IsNullOrWhiteSpace(sampleColor) ? "(blank)" : sampleColor) + "\n\n"
            + "Expected examples: #1F84BB, #D3CEC2, #80FF0000.\n\n"
            + "This may still be valid if the measure only returns a color inside a visual filter context. Continue anyway?",
            "Validate color measure",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);
        if (response != DialogResult.Yes) { Info("Cancelled - selected color measure did not validate as a HEX color."); return; }
    }
}
catch (Exception ex)
{
    var response = MessageBox.Show(
        "Could not evaluate the selected color measure before editing PBIR.\n\n"
        + "Measure: " + colorMeasure.DaxObjectFullName + "\n"
        + "Error: " + ex.Message + "\n\n"
        + "Continue anyway?",
        "Validate color measure",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning);
    if (response != DialogResult.Yes) { Info("Cancelled - color measure validation could not run."); return; }
}

// ---------------------------------------------------------------------------
// Step 3: build the flat leaf catalog (Page > VisualType > Visual >
// Element ["Background"/"Font"] > [by CategoryField >] Field), then let the
// user check as many targets as they like in one multi-select tree.
// ---------------------------------------------------------------------------
Func<string, string, string, string, string, bool, string, string, string> makeKey =
    (file, container, propKey, fieldEntity, fieldProperty, detail, catEntity, catProperty) =>
        string.Join("||", file, container, propKey, fieldEntity, fieldProperty, detail.ToString(), catEntity ?? "", catProperty ?? "");

Func<string, int?, bool, JObject> buildSelector = (metaKey, matchingOption, defaultSelector) =>
{
    if (defaultSelector)
        return new JObject { ["id"] = "default" };

    var selector = new JObject { ["metadata"] = metaKey };
    if (matchingOption.HasValue)
        selector["data"] = new JArray(new JObject
        {
            ["dataViewWildcard"] = new JObject { ["matchingOption"] = matchingOption.Value }
        });
    return selector;
};

Func<JObject, string, string, int?, bool, JObject> findSelectorTemplate = (root, containerName, metaKey, matchingOption, defaultSelector) =>
{
    var arr = root["visual"]?["objects"]?[containerName] as JArray;
    if (arr != null)
    {
        foreach (var entry in arr.OfType<JObject>())
        {
            var selector = entry["selector"] as JObject;
            if (selector == null) continue;

            if (defaultSelector)
            {
                if ((string)selector["id"] == "default")
                    return (JObject)selector.DeepClone();
                continue;
            }

            if ((string)selector["metadata"] != metaKey) continue;
            if (matchingOption.HasValue)
            {
                var existingOption = (int?)selector["data"]?[0]?["dataViewWildcard"]?["matchingOption"];
                if (existingOption != matchingOption.Value) continue;
            }
            return (JObject)selector.DeepClone();
        }
    }
    return buildSelector(metaKey, matchingOption, defaultSelector);
};

var allLeaves = new List<LeafConfig>();

foreach (var v in allVisuals)
{
    List<(string Label, string PropertyKey, string ContainerName, int? MatchingOption, bool DefaultSelector)> elements;
    if (tableMatrixTypes.Contains(v.VisualType))
        elements = new List<(string, string, string, int?, bool)>
        {
            ("Background", "backColor", "values", 1, false),
            ("Font", "fontColor", "values", 1, false)
        };
    else if (barColumnTypes.Contains(v.VisualType)) // already filtered to exactly one value field
        elements = new List<(string, string, string, int?, bool)> { ("Background", "fill", "dataPoint", 1, false) };
    else // advanced card / cardVisual
        elements = new List<(string, string, string, int?, bool)>
        {
            ("Card fill", "fillColor", "fillCustom", null, true),
            ("Outline color", "lineColor", "outline", null, true),
            ("Shadow color", "color", "shadowCustom", null, true),
            ("Value font", "fontColor", "value", 0, false),
            ("Category label font", "fontColor", "label", 0, false),
            ("Reference detail font", "detailFontColor", "referenceLabelDetail", 0, false),
            ("Reference detail background", "detailBackgroundColor", "referenceLabelDetail", 0, false),
            ("Reference value font", "valueFontColor", "referenceLabelValue", 0, false),
            ("Reference title font", "titleFontColor", "referenceLabelTitle", 0, false)
        };

    foreach (var el in elements)
    {
        var targetFields = el.DefaultSelector ? v.ValueFields.Take(1) : v.ValueFields;
        foreach (var vf in targetFields)
        {
            var metaKey = vf.Entity + "." + vf.Property;
            allLeaves.Add(new LeafConfig
            {
                File = v.File,
                Root = v.Root,
                ContainerName = el.ContainerName,
                PropertyKey = el.PropertyKey,
                FieldEntity = vf.Entity,
                FieldProperty = vf.Property,
                MatchingOption = el.MatchingOption,
                DefaultSelector = el.DefaultSelector,
                SelectorTemplate = findSelectorTemplate(v.Root, el.ContainerName, metaKey, el.MatchingOption, el.DefaultSelector),
                PageDisplay = v.PageDisplay,
                VisualType = v.VisualType,
                VisualLabel = v.Label,
                ElementLabel = el.Label,
                DisplayPath = v.PageDisplay + " > " + v.VisualType + " > " + v.Label + " > " + el.Label + (el.DefaultSelector ? " > Visual" : " > " + vf.Entity + "." + vf.Property),
                Key = makeKey(v.File, el.ContainerName, el.PropertyKey, vf.Entity, vf.Property, el.DefaultSelector, el.MatchingOption.HasValue ? el.MatchingOption.Value.ToString() : null, null)
            });
        }
    }
}
if (allLeaves.Count == 0) Error("No formattable Background/Font targets found in this report.");

var selectedLeaves = pickManyLeaves(allLeaves, "Pick every target to apply '" + colorMeasure.Name + "' to");
if (cancelled || selectedLeaves == null) { Info("Cancelled - no files were changed."); return; }

// ---------------------------------------------------------------------------
// Step 4: preview + confirm
// ---------------------------------------------------------------------------
var dt = new System.Data.DataTable();
dt.Columns.Add("Target");
dt.Columns.Add("Property");
dt.Columns.Add("Scope");
dt.Columns.Add("Color measure");
foreach (var leaf in selectedLeaves)
    dt.Rows.Add(
        leaf.DisplayPath,
        leaf.PropertyKey,
        leaf.DefaultSelector ? "Whole visual" : "Whole column / whole field",
        colorMeasure.DaxObjectFullName);
Output(dt);

var confirm = MessageBox.Show(
    "Apply " + selectedLeaves.Count + " conditional formatting " + (selectedLeaves.Count == 1 ? "entry" : "entries") + "?\n\n"
    + "Make sure Power BI Desktop is closed and your changes are committed to git before continuing.",
    "Confirm changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
if (confirm != DialogResult.Yes) { Info("Cancelled - no files were changed."); return; }

// ---------------------------------------------------------------------------
// Step 5: write all changes
// ---------------------------------------------------------------------------
Func<JObject> makeColorExpr = () => JObject.Parse(
    "{ \"expr\": { \"Measure\": { \"Expression\": { \"SourceRef\": { \"Entity\": \"" + colorMeasure.Table.Name + "\" } }, \"Property\": \"" + colorMeasure.Name + "\" } } }");

var touchedFiles = new Dictionary<string, JObject>();

foreach (var leaf in selectedLeaves)
{
    var visualObj = leaf.Root["visual"] as JObject;
    var objects = visualObj["objects"] as JObject;
    if (objects == null) { objects = new JObject(); visualObj["objects"] = objects; }

    var arr = objects[leaf.ContainerName] as JArray;
    if (arr == null) { arr = new JArray(); objects[leaf.ContainerName] = arr; }

    string metaKey = leaf.FieldEntity + "." + leaf.FieldProperty;

    JObject selector = leaf.SelectorTemplate != null
        ? (JObject)leaf.SelectorTemplate.DeepClone()
        : buildSelector(metaKey, leaf.MatchingOption, leaf.DefaultSelector);

    Func<JObject, bool> matchesEntry = (e) =>
    {
        var sel = e["selector"] as JObject;
        if (sel == null) return false;

        if (leaf.DefaultSelector)
            return (string)sel["id"] == "default";

        if ((string)sel["metadata"] != metaKey) return false;
        if (sel["data"]?[0]?["scopeId"] != null) return false;
        if (leaf.MatchingOption.HasValue)
            return (int?)sel["data"]?[0]?["dataViewWildcard"]?["matchingOption"] == leaf.MatchingOption.Value;
        return true;
    };

    Action<JObject, Func<JObject, bool>> upsertEntry = (entrySelector, matcher) =>
    {
        JObject entry = arr.OfType<JObject>().FirstOrDefault(matcher);
        if (entry == null)
        {
            entry = new JObject { ["properties"] = new JObject(), ["selector"] = entrySelector };
            arr.Add(entry);
        }

        var props = entry["properties"] as JObject;
        if (props == null) { props = new JObject(); entry["properties"] = props; }
        props[leaf.PropertyKey] = new JObject { ["solid"] = new JObject { ["color"] = makeColorExpr() } };
    };

    upsertEntry(selector, matchesEntry);

    if (barColumnTypes.Contains(leaf.VisualType) && leaf.ContainerName == "dataPoint" && leaf.PropertyKey == "fill")
    {
        JObject wildcardSelector = JObject.Parse("{ \"data\": [ { \"dataViewWildcard\": { \"matchingOption\": 1 } } ] }");
        upsertEntry(wildcardSelector, e =>
            e["selector"]?["metadata"] == null &&
            (int?)e["selector"]?["data"]?[0]?["dataViewWildcard"]?["matchingOption"] == 1);
    }

    touchedFiles[leaf.File] = leaf.Root;
}

foreach (var kv in touchedFiles)
    File.WriteAllText(kv.Key, kv.Value.ToString(Newtonsoft.Json.Formatting.Indented));

Info("Done. Applied '" + colorMeasure.DaxObjectFullName + "' to " + selectedLeaves.Count + " target(s) across " + touchedFiles.Count + " visual.json file(s)."
    + "\n\nRemember: reopen Power BI Desktop to see the change (it does not pick up PBIR edits live).");
