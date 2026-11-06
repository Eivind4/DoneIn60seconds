// ADD RANKING AS DAX UDFs
// =========================
//
// Purpose: Generate reusable, model-DEPENDENT "Local.Ranking.*" DAX User Defined Functions (UDFs)
// for ranking a base measure by one or more selected columns, based on the ranking patterns from:
//   https://www.daxpatterns.com/ranking/
//
// Column selection is limited to columns from tables that behave like ranking dimensions: not a
// disconnected date-table (no active relationships), not a calculation group, and not a "fact"
// table (i.e. the table must filter at least one other table via an active relationship).
//
// For each selected column, up to five functions are created:
//   - Local.Ranking.<Column>(baseMeasure)                   - the dynamic rank position of
//         baseMeasure within the selected column (RANKX over ALLSELECTED(column), blank outside
//         the column's scope or when baseMeasure is blank). Single-parameter, so it can be
//         wrapped into a measure with "5 - Create measure from model dependent functions.cs".
//   - Local.Ranking.<Column>TopN(baseMeasure)                - baseMeasure itself, but only shown
//         when its rank (via the function above) is within the Top N (else blank) - e.g.
//         "profitability for the 10 customers with highest sales". Single-parameter.
//   - Local.Ranking.<Column>TopNbyOtherMeasure(rankMeasure, summarizeMeasure) - same Top N logic,
//         but ranks by one measure and returns/summarizes a DIFFERENT measure - e.g. rank by Sales
//         Amount, return Profit Margin for the top N. Two-parameter, so it is NOT wrapped by the
//         single-parameter macro above - instead it is picked up automatically by "5 - Compare
//         first selected measure with second selected measure.cs" (any function with exactly two
//         measure-compatible parameters is offered there).
//   - Local.Ranking.<Column>InCategory(baseMeasure) (optional - only if the "Ranking in Category"
//         option is checked, a category column is picked, and exactly ONE ranking column was
//         selected, since the category column must be on the SAME table) - rank restricted to
//         rows sharing the same category column value (daxpatterns.com "Rank in Category"
//         pattern). Single-parameter.
//   - Local.Ranking.Rank_TopN_<Column>_with_others(baseMeasure) (optional - only if "Ranking
//         with Others" is checked). It creates a ranking table named "<Table> <Column> Rank",
//         relates it one-to-many to the selected dimension column, and returns baseMeasure for
//         the Top N group or the aggregated remainder for the Others group.
//
// The Top N value is either a fixed literal (hard-coded, e.g. 10), or dynamic via a "Top N" /
// "Top N Value" calculated table + measure (same convention as "4 - Create measures TopN based on
// selected column(s) and measure(s).cs"), with a default range of 1-20 and increment 1.
//
// Annotations added to the single-parameter Local functions (read by "5 - Create measure from
// model dependent functions.cs"):
//   - FolderSuffix  = "Ranking" (all single-parameter variants)
//   - MeasurePrefix = "Ranking <Column> by"                       (rank position)
//                     "Ranking <Column> in <CategoryColumn> by"   (in-category rank position)
//                     "TopN" (dynamic) / "Top<N>" (fixed, e.g. "Top10")   (Top N variant)
//   - MeasureSuffix = "by <Column>"                               (Top N variant only)
//                     "by <Column> with <Others label>"           (Top N with Others variant)
//   - FormatString  = "#,0" (rank position and in-category rank position only - a whole rank
//                     number; the Top N/TopNbyOtherMeasure variants inherit the base measure's
//                     own format, since they return that measure's value, not a rank).
// E.g. wrapping base measure [Sum Amount] gives "Ranking Customer by Sum Amount", "TopN Sum Amount
// by Customer" (or "Top10 Sum Amount by Customer" if fixed), and "Ranking Customer in Customer
// Group by Sum Amount".
// The two-parameter TopNbyOtherMeasure function has no MeasurePrefix/MeasureSuffix annotations,
// since it is consumed via the Compare-measures macro instead.
//
// Safety checks performed before creating each function:
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function is reused/referenced instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one (and adjust the
//       MeasurePrefix/MeasureSuffix annotations), Overwrite the existing (outdated) definition, or
//       Skip creating it (never silently overwritten).
//
#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;

Application.UseWaitCursor = false;
WaitFormVisible = false;

Func<string, string> NormalizeExpr = s => string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
Func<string, string> SanitizeIdentifier = s => new string((s ?? "").Where(char.IsLetterOrDigit).ToArray());

// ================================================================================================
// DIALOG 1: Fixed vs Dynamic Top N input.
// ================================================================================================
var topNForm = new Form
{
    Text = "Ranking - Top N input",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 480,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var topNFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
topNForm.Controls.Add(topNFlow);
Action<string> topNLabel = s => topNFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

topNLabel("Choose how the Top N value is determined:");
var fixedRadio   = new RadioButton { Text = "Fixed (hard-coded N)", AutoSize = true, Checked = true };
var dynamicRadio = new RadioButton { Text = "Dynamic (adds a 'Top N' calculated table with a slicer)", AutoSize = true };
topNFlow.Controls.Add(fixedRadio);
topNFlow.Controls.Add(dynamicRadio);

// -- Fixed panel --
var fixedPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0) };
fixedPanel.Controls.Add(new Label { Text = "N (top count for standard ranking):", AutoSize = true });
var nFixedInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 100000, Value = 10 };
fixedPanel.Controls.Add(nFixedInput);
topNFlow.Controls.Add(fixedPanel);

// -- Dynamic panel --
var dynamicPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(20, 5, 0, 0), Visible = false };
dynamicPanel.Controls.Add(new Label { Text = "Table minimum:", AutoSize = true });
var minDynInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 100000, Value = 1 };
dynamicPanel.Controls.Add(minDynInput);
dynamicPanel.Controls.Add(new Label { Text = "Table maximum:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var maxDynInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 100000, Value = 20 };
dynamicPanel.Controls.Add(maxDynInput);
dynamicPanel.Controls.Add(new Label { Text = "Increment:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
var incDynInput = new NumericUpDown { Width = 120, Minimum = 1, Maximum = 10000, Value = 1 };
dynamicPanel.Controls.Add(incDynInput);
topNFlow.Controls.Add(dynamicPanel);

Action updateTopNVis = () => { fixedPanel.Visible = fixedRadio.Checked; dynamicPanel.Visible = dynamicRadio.Checked; };
fixedRadio.CheckedChanged   += (s, e) => updateTopNVis();
dynamicRadio.CheckedChanged += (s, e) => updateTopNVis();

var topNBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 15, 0, 0) };
var topNOk     = new Button { Text = "OK",     AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var topNCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
topNBtnPanel.Controls.Add(topNOk);
topNBtnPanel.Controls.Add(topNCancel);
topNFlow.Controls.Add(topNBtnPanel);
topNForm.AcceptButton = topNOk;
topNForm.CancelButton = topNCancel;

if (topNForm.ShowDialog() == DialogResult.Cancel) return;

bool isDynamicTopN = dynamicRadio.Checked;
var notes = new List<string>();
string topNArg;
// Used as the MeasurePrefix for the TopN function's wrapped measures, e.g. "TopN Sum Amount by Customer"
// (dynamic) or "Top10 Sum Amount by Customer" (fixed, N baked into the label).
string topNPrefixLabel = isDynamicTopN ? "TopN" : $"Top{(int)nFixedInput.Value}";

if (!isDynamicTopN)
{
    topNArg = ((int)nFixedInput.Value).ToString();
}
else
{
    const string topNTableName = "Top N";
    const string topNMeasureName = "Top N Value";
    var existingTable = Model.Tables.FirstOrDefault(t => t.Name == topNTableName);

    bool createTableAndMeasure = true;
    int fallbackDefault = (int)nFixedInput.Value;

    if (existingTable != null)
    {
        var choice = MessageBox.Show(
            $"A table named '{topNTableName}' already exists.\n\n" +
            "Yes = Overwrite (delete and recreate it with the values above)\n" +
            "No = Reuse the existing table as-is (values above are ignored)\n" +
            "Cancel = Abort this script",
            "Top N table already exists",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (choice == DialogResult.Cancel) return;

        if (choice == DialogResult.Yes)
        {
            existingTable.Delete();
            existingTable = null;
        }
        else
        {
            createTableAndMeasure = false;
        }
    }

    if (createTableAndMeasure)
    {
        int min = (int)minDynInput.Value;
        int max = (int)maxDynInput.Value;
        int inc = (int)incDynInput.Value;

        var table = Model.AddCalculatedTable(topNTableName, $"GENERATESERIES({min}, {max}, {inc})");
        table.TableGroup = "Parameters";
        table.Description = "Selected Top N value used by Ranking DaxUDFs";

        var valueColumn = table.Columns.Contains("Value") ? table.Columns["Value"] : table.Columns.First();
        valueColumn.SummarizeBy = AggregateFunction.None;
        valueColumn.Description = "Selected parameter for Top N";
        valueColumn.FormatString = "#,0";

        var valueMeasure = table.AddMeasure(topNMeasureName, $"SELECTEDVALUE({valueColumn.DaxObjectFullName}, {fallbackDefault})");
        valueMeasure.Description = $"Selected Top N value, defaults to {fallbackDefault} if nothing selected";
        valueMeasure.FormatString = "#,0";

        existingTable = table;
    }

    var topNMeasureRef = existingTable.Measures.FirstOrDefault(m => m.Name == topNMeasureName);
    if (topNMeasureRef == null)
    {
        Error($"Could not find measure '{topNMeasureName}' on the existing '{topNTableName}' table. Re-run and choose Overwrite.");
        return;
    }
    topNArg = topNMeasureRef.DaxObjectFullName;
}

// ================================================================================================
// DIALOG 2: Select ranking column(s) (multi-select, search + tree grouped by Table > Folder path >
// Column), plus an optional single "category" column for the "Ranking in Category" variant.
// ================================================================================================

// Excludes: disconnected date-tables (no active relationships), calculation group tables, and
// "fact" tables (tables that never filter another table via an active relationship - i.e. they
// are only ever on the "many" side, so ranking by one of their columns would not behave like a
// dimension). A table qualifies as a valid ranking dimension if it is on the filtering ("one")
// side of at least one active relationship.
Func<Table, bool> HasActiveRelationship = t => Model.Relationships.Any(r => r.IsActive && (r.FromTable == t || r.ToTable == t));
Func<Table, bool> FiltersAnotherTable = t => Model.Relationships.Any(r => r.IsActive && r.ToTable == t);
Func<Table, bool> IsDateTable = t => Model.AllCalendars.Any(cal => t.Calendars.Contains(cal)) || t.Columns.Any(c => c.IsKey && c.DataType == DataType.DateTime);
Func<Table, bool> IsDisconnectedDateTable = t => IsDateTable(t) && !HasActiveRelationship(t);
Func<Table, bool> IsCalcGroupTable = t => t is CalculationGroupTable;

var allColumns = Model.AllColumns
    .Where(c => c.Table.Name != "Top N" && c.Table.Name != "Conditional Format Input")
    .Where(c => !IsCalcGroupTable(c.Table))
    .Where(c => !IsDisconnectedDateTable(c.Table))
    .Where(c => FiltersAnotherTable(c.Table))
    .ToList();

if (allColumns.Count == 0)
{
    Info("No suitable dimension columns found in the model (after excluding disconnected date-tables, calculation groups, and fact tables).");
    return;
}

// (Re)builds a column-picker TreeView, grouped by Table > DisplayFolder segments > Column name,
// filtered live by the given search text. If multiSelect is true, checkboxes behave as a
// multi-select set (tracked in pickedSet); otherwise checking a leaf unchecks all other leaves
// (single-select), tracked via getCurrentSelection/setSelected.
Action<TreeView, string, List<Column>, HashSet<Column>> RebuildColumnTree = (tree, filterText, columns, pickedSet) =>
{
    tree.BeginUpdate();
    tree.Nodes.Clear();

    var filtered = columns
        .Where(c => string.IsNullOrWhiteSpace(filterText) ||
                    c.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.Table.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
        .OrderBy(c => c.Table.Name).ThenBy(c => c.DisplayFolder).ThenBy(c => c.Name)
        .ToList();

    var nodeLookup = new Dictionary<string, TreeNode>();

    foreach (var c in filtered)
    {
        var pathParts = new List<string> { c.Table.Name };
        if (!string.IsNullOrWhiteSpace(c.DisplayFolder))
            pathParts.AddRange(c.DisplayFolder.Split('\\').Where(s => !string.IsNullOrWhiteSpace(s)));

        TreeNodeCollection currentLevel = tree.Nodes;
        string pathSoFar = "";
        foreach (var part in pathParts)
        {
            pathSoFar = string.IsNullOrEmpty(pathSoFar) ? part : pathSoFar + "\\" + part;
            TreeNode node;
            if (!nodeLookup.TryGetValue(pathSoFar, out node))
            {
                node = new TreeNode(part);
                currentLevel.Add(node);
                nodeLookup[pathSoFar] = node;
            }
            currentLevel = node.Nodes;
        }

        var leaf = new TreeNode(c.Name) { Tag = c, Checked = pickedSet != null && pickedSet.Contains(c) };
        currentLevel.Add(leaf);
    }

    tree.ExpandAll();
    tree.EndUpdate();
};

var pickForm = new Form
{
    Text = "Select ranking column(s)",
    Width = 1000,
    Height = Math.Min(850, Screen.PrimaryScreen.WorkingArea.Height - 80),
    StartPosition = FormStartPosition.CenterScreen,
    FormBorderStyle = FormBorderStyle.Sizable,
    MinimizeBox = false,
    AutoScroll = true,
    AutoScrollMinSize = new Size(960, 790)
};

Func<string, int, int, int, Label> BuildLabel = (text, x, y, w) =>
    new Label { Text = text, Location = new Point(x, y), Width = w, Height = 20, Font = new Font(pickForm.Font, FontStyle.Bold) };

// -- Left: ranking columns (multi-select) --
var rankLabel  = BuildLabel("Ranking column(s) - baseMeasure will be ranked by these:", 20, 15, 440);
var rankSearch = new TextBox { Location = new Point(20, 40), Width = 440 };
var rankTree   = new TreeView { Location = new Point(20, 68), Width = 440, Height = 430, CheckBoxes = true };
var pickedColumns = new HashSet<Column>();

// -- Right: optional category column (single-select) --
var categoryCheckbox = new CheckBox { Text = "Also create a 'Ranking in Category' variant", Location = new Point(500, 15), Width = 440, AutoSize = false, Height = 20 };
var categoryHint      = new Label { Text = "(select exactly one ranking column on the left first)", Location = new Point(500, 40), Width = 440, AutoSize = false, Height = 16, ForeColor = Color.DimGray, Visible = false };
var categorySearch   = new TextBox { Location = new Point(500, 40), Width = 440, Enabled = false };
var categoryTree     = new TreeView { Location = new Point(500, 68), Width = 440, Height = 430, CheckBoxes = true, Enabled = false };
Column selectedCategoryColumn = null;

// Category-variant columns must be on the SAME table as the (single) ranking column, and cannot
// be the ranking column itself - enforced here at the picker level (rather than only failing
// later at UDF-generation time) so an invalid combination can never be selected in the first place.
Func<List<Column>> CategoryEligibleColumns = () =>
    pickedColumns.Count == 1
        ? allColumns.Where(c => c.Table == pickedColumns.First().Table && c != pickedColumns.First()).ToList()
        : new List<Column>();

Action RefreshCategoryTree = null;

bool suppressRank = false;
Action<TreeNodeCollection> ToggleRankDescendants = null;
ToggleRankDescendants = nodes =>
{
    foreach (TreeNode n in nodes)
    {
        var col = n.Tag as Column;
        if (col != null)
        {
            if (n.Checked) pickedColumns.Add(col); else pickedColumns.Remove(col);
        }
        ToggleRankDescendants(n.Nodes);
    }
};
rankTree.AfterCheck += (s, e) =>
{
    if (suppressRank) return;
    suppressRank = true;
    var col = e.Node.Tag as Column;
    if (col != null)
    {
        if (e.Node.Checked) pickedColumns.Add(col); else pickedColumns.Remove(col);
    }
    else
    {
        foreach (TreeNode child in e.Node.Nodes) child.Checked = e.Node.Checked;
        ToggleRankDescendants(e.Node.Nodes);
    }
    suppressRank = false;
    RefreshCategoryTree();
};

bool suppressCategory = false;
Action<TreeNodeCollection> UncheckAllCategory = null;
UncheckAllCategory = nodes =>
{
    foreach (TreeNode n in nodes)
    {
        if (n.Checked) n.Checked = false;
        UncheckAllCategory(n.Nodes);
    }
};
categoryTree.AfterCheck += (s, e) =>
{
    if (suppressCategory) return;
    var col = e.Node.Tag as Column;
    if (col == null) { e.Node.Checked = false; return; } // group nodes are not selectable (single-select)
    suppressCategory = true;
    if (e.Node.Checked)
    {
        UncheckAllCategory(categoryTree.Nodes);
        e.Node.Checked = true;
        selectedCategoryColumn = col;
    }
    else
    {
        selectedCategoryColumn = null;
    }
    suppressCategory = false;
};

RebuildColumnTree(rankTree, "", allColumns, pickedColumns);

// Rebuilds the category tree from CategoryEligibleColumns() (same table as the single picked
// ranking column, excluding that column), clearing the current selection if it is no longer
// eligible (e.g. the ranking column selection changed), and shows/hides the "pick one ranking
// column first" hint accordingly.
RefreshCategoryTree = () =>
{
    var eligible = CategoryEligibleColumns();
    if (selectedCategoryColumn != null && !eligible.Contains(selectedCategoryColumn))
        selectedCategoryColumn = null;

    bool hasEligible = eligible.Count > 0;
    categoryHint.Visible = !hasEligible;
    categorySearch.Visible = hasEligible;
    categorySearch.Enabled = hasEligible && categoryCheckbox.Checked;
    categoryTree.Enabled = hasEligible && categoryCheckbox.Checked;

    RebuildColumnTree(categoryTree, categorySearch.Text, eligible,
        selectedCategoryColumn == null ? null : new HashSet<Column> { selectedCategoryColumn });
};
RefreshCategoryTree();

rankSearch.TextChanged += (s, e) => RebuildColumnTree(rankTree, rankSearch.Text, allColumns, pickedColumns);
categorySearch.TextChanged += (s, e) => RebuildColumnTree(categoryTree, categorySearch.Text, CategoryEligibleColumns(),
    selectedCategoryColumn == null ? null : new HashSet<Column> { selectedCategoryColumn });

categoryCheckbox.CheckedChanged += (s, e) => RefreshCategoryTree();

// -- Optional Ranking with Others variant --
var othersCheckbox = new CheckBox
{
    Text = "Also create a 'Ranking with Others' variant and ranking table(s)",
    Location = new Point(20, 515),
    Width = 600,
    AutoSize = false,
    Height = 20
};
var othersOptions = new FlowLayoutPanel
{
    Location = new Point(40, 540),
    Width = 900,
    Height = 145,
    FlowDirection = FlowDirection.TopDown,
    WrapContents = false,
    Enabled = false
};
othersOptions.Controls.Add(new Label
{
    Text = "The ranking table uses '<Dimension table> <Column> Rank', with a one-to-many relationship to the selected dimension column.",
    AutoSize = true
});
var calculatedTableRadio = new RadioButton { Text = "Create calculated table", AutoSize = true, Checked = true };
var powerQueryTableRadio = new RadioButton { Text = "Create Import table in Power Query", AutoSize = true };
var tableTypeRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
tableTypeRow.Controls.Add(calculatedTableRadio);
tableTypeRow.Controls.Add(powerQueryTableRadio);
othersOptions.Controls.Add(tableTypeRow);
var rankingGroupRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
rankingGroupRow.Controls.Add(new Label { Text = "Ranking group prefix:", AutoSize = true, Margin = new Padding(0, 4, 8, 0) });
var rankingGroupPrefixInput = new TextBox { Width = 180, Text = "Best" };
rankingGroupRow.Controls.Add(rankingGroupPrefixInput);
rankingGroupRow.Controls.Add(new Label { Text = "(creates e.g. 'Best Product'; enter 'Worst' when applicable)", AutoSize = true, Margin = new Padding(8, 4, 0, 0) });
othersOptions.Controls.Add(rankingGroupRow);
var othersNameRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
othersNameRow.Controls.Add(new Label { Text = "Others name and group:", AutoSize = true, Margin = new Padding(0, 4, 8, 0) });
var othersNameInput = new TextBox { Width = 180, Text = "Others" };
othersNameRow.Controls.Add(othersNameInput);
othersNameRow.Controls.Add(new Label { Text = "(change this if the source category already contains 'Others')", AutoSize = true, Margin = new Padding(8, 4, 0, 0) });
othersOptions.Controls.Add(othersNameRow);

othersCheckbox.CheckedChanged += (s, e) => othersOptions.Enabled = othersCheckbox.Checked;

var pickBtnPanel = new FlowLayoutPanel { Location = new Point(20, 710), AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
var pickOk     = new Button { Text = "OK",     AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
var pickCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
pickBtnPanel.Controls.Add(pickOk);
pickBtnPanel.Controls.Add(pickCancel);

// The "Ranking in Category" variant needs a single dimension in the SAME table as the ranking
// column, so it only makes sense when exactly one ranking column is selected - validated here
// instead of via DialogResult, so the dialog stays open on validation failure.
pickOk.Click += (s, e) =>
{
    if (categoryCheckbox.Checked && pickedColumns.Count != 1)
    {
        MessageBox.Show(
            "The 'Ranking in Category' variant requires exactly one ranking column to be selected.\n\nEither uncheck it, or select only one ranking column.",
            "Select exactly one ranking column",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
    }
    if (othersCheckbox.Checked && string.IsNullOrWhiteSpace(rankingGroupPrefixInput.Text))
    {
        MessageBox.Show(
            "Enter a ranking group prefix, such as 'Best' or 'Worst'.",
            "Ranking group prefix is required",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
    }
    if (othersCheckbox.Checked && string.IsNullOrWhiteSpace(othersNameInput.Text))
    {
        MessageBox.Show(
            "Enter the label used for the Others ranking name and ranking group.",
            "Others label is required",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
    }
    pickForm.DialogResult = DialogResult.OK;
    pickForm.Close();
};

pickForm.Controls.Add(rankLabel);
pickForm.Controls.Add(rankSearch);
pickForm.Controls.Add(rankTree);
pickForm.Controls.Add(categoryCheckbox);
pickForm.Controls.Add(categoryHint);
pickForm.Controls.Add(categorySearch);
pickForm.Controls.Add(categoryTree);
pickForm.Controls.Add(othersCheckbox);
pickForm.Controls.Add(othersOptions);
pickForm.Controls.Add(pickBtnPanel);
pickForm.AcceptButton = pickOk;
pickForm.CancelButton = pickCancel;

if (pickForm.ShowDialog() == DialogResult.Cancel) return;

bool wantCategoryVariant = categoryCheckbox.Checked;
bool wantOthersVariant = othersCheckbox.Checked;
bool createOthersCalculatedTable = calculatedTableRadio.Checked;
string rankingGroupPrefix = rankingGroupPrefixInput.Text.Trim();
string othersName = othersNameInput.Text.Trim();
if (wantCategoryVariant && selectedCategoryColumn == null)
{
    notes.Add("'Ranking in Category' was requested but no category column was selected - the category variant was skipped for all columns.");
    wantCategoryVariant = false;
}

if (pickedColumns.Count == 0)
{
    Info("No ranking column(s) selected. Nothing to do.");
    return;
}

// A separate, fixed Top N can be used for the Ranking in Category UDF. Dynamic Top N uses the
// shared slicer measure, so the category and standard variants always respond to the same slicer.
string topNCategoryArg = topNArg;
if (wantCategoryVariant && !isDynamicTopN)
{
    var categoryTopNForm = new Form
    {
        Text = "Ranking in Category - Top N input",
        MaximizeBox = false,
        MinimizeBox = false,
        ShowIcon = false,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        Width = 360,
        FormBorderStyle = FormBorderStyle.FixedDialog,
        StartPosition = FormStartPosition.CenterScreen
    };
    var categoryTopNFlow = new FlowLayoutPanel
    {
        Padding = new Padding(10),
        FlowDirection = FlowDirection.TopDown,
        WrapContents = false,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink
    };
    categoryTopNForm.Controls.Add(categoryTopNFlow);
    categoryTopNFlow.Controls.Add(new Label
    {
        Text = "N (top count within each category):",
        AutoSize = true
    });
    var categoryTopNInput = new NumericUpDown
    {
        Width = 120,
        Minimum = 1,
        Maximum = 100000,
        Value = 3
    };
    categoryTopNFlow.Controls.Add(categoryTopNInput);
    var categoryTopNButtons = new FlowLayoutPanel
    {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        FlowDirection = FlowDirection.LeftToRight,
        Margin = new Padding(0, 15, 0, 0)
    };
    var categoryTopNOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
    var categoryTopNCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    categoryTopNButtons.Controls.Add(categoryTopNOk);
    categoryTopNButtons.Controls.Add(categoryTopNCancel);
    categoryTopNFlow.Controls.Add(categoryTopNButtons);
    categoryTopNForm.AcceptButton = categoryTopNOk;
    categoryTopNForm.CancelButton = categoryTopNCancel;

    if (categoryTopNForm.ShowDialog() == DialogResult.Cancel) return;
    topNCategoryArg = ((int)categoryTopNInput.Value).ToString();
}

// ================================================================================================
// Create the Local.Ranking.* functions - same skip/rename/overwrite conflict handling used by the
// other DaxUDF generator macros. Supports both MeasurePrefix and MeasureSuffix (pass "" for
// whichever does not apply), plus an optional FormatString annotation (e.g. "#,0" for rank
// position functions, which return a whole-number rank rather than the base measure's own format).
// ================================================================================================
Func<string, string, string, string, string, string, string, string> CreateLocalUDF = (funcName, daxBody, description, measurePrefix, measureSuffix, folderSuffix, formatString) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        notes.Add($"'{funcName}' was NOT created - using the existing function '{dupExprFunc.Name}' instead, since it has an identical expression.");
        return dupExprFunc.Name;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            notes.Add($"'{funcName}' already exists with the same expression - left unchanged.");
            return existing.Name;
        }

        var conflictDlg = new Form { Text = "Function name already exists", Width = 500, Height = 340, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false };
        var conflictFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = $"A function named '{funcName}' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(460, 0), Margin = new Padding(0, 0, 0, 10) });

        var renameRow   = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var renameRadio = new RadioButton { Text = "Rename to:", AutoSize = true, Checked = true, Margin = new Padding(0, 3, 8, 0) };
        var renameBox   = new TextBox { Width = 250, Text = funcName };
        renameRow.Controls.Add(renameRadio);
        renameRow.Controls.Add(renameBox);
        conflictFlow.Controls.Add(renameRow);

        var prefixRow  = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(20, 0, 0, 5) };
        prefixRow.Controls.Add(new Label { Text = "MeasurePrefix annotation to use:", AutoSize = true, Margin = new Padding(0, 3, 8, 0) });
        var prefixBox = new TextBox { Width = 200, Text = measurePrefix };
        prefixRow.Controls.Add(prefixBox);
        conflictFlow.Controls.Add(prefixRow);

        var suffixRow  = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(20, 0, 0, 5) };
        suffixRow.Controls.Add(new Label { Text = "MeasureSuffix annotation to use:", AutoSize = true, Margin = new Padding(0, 3, 8, 0) });
        var suffixBox = new TextBox { Width = 200, Text = measureSuffix };
        suffixRow.Controls.Add(suffixBox);
        conflictFlow.Controls.Add(suffixRow);

        var overwriteRadio = new RadioButton { Text = "Overwrite the existing definition with this new one", AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var skipRadio       = new RadioButton { Text = "Skip - do not create this UDF", AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        conflictFlow.Controls.Add(overwriteRadio);
        conflictFlow.Controls.Add(skipRadio);

        var conflictBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
        var conflictOk = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
        conflictBtnPanel.Controls.Add(conflictOk);
        conflictFlow.Controls.Add(conflictBtnPanel);
        conflictDlg.AcceptButton = conflictOk;

        conflictDlg.ShowDialog();
        measurePrefix = string.IsNullOrWhiteSpace(prefixBox.Text) ? measurePrefix : prefixBox.Text.Trim();
        measureSuffix = string.IsNullOrWhiteSpace(suffixBox.Text) ? measureSuffix : suffixBox.Text.Trim();

        if (overwriteRadio.Checked)
        {
            existing.Expression  = daxBody;
            existing.Description = description;
            existing.SetAnnotation("FolderSuffix",  folderSuffix);
            if (!string.IsNullOrWhiteSpace(measurePrefix)) existing.SetAnnotation("MeasurePrefix", measurePrefix);
            if (!string.IsNullOrWhiteSpace(measureSuffix)) existing.SetAnnotation("MeasureSuffix", measureSuffix);
            if (!string.IsNullOrWhiteSpace(formatString))  existing.SetAnnotation("FormatString",  formatString);
            notes.Add($"'{funcName}' already existed with a different expression - its definition was OVERWRITTEN with the new one.");
            return existing.Name;
        }
        else if (skipRadio.Checked)
        {
            notes.Add($"'{funcName}' was NOT created - a function with that name already exists (existing definition kept).");
            return null;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                notes.Add($"'{funcName}' was NOT created - no valid new name was provided (existing definition kept).");
                return null;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                notes.Add($"'{funcName}' was NOT created - the alternate name '{newName}' also already exists.");
                return null;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
    fn.SetAnnotation("FolderSuffix", folderSuffix);
    if (!string.IsNullOrWhiteSpace(measurePrefix)) fn.SetAnnotation("MeasurePrefix", measurePrefix);
    if (!string.IsNullOrWhiteSpace(measureSuffix)) fn.SetAnnotation("MeasureSuffix", measureSuffix);
    if (!string.IsNullOrWhiteSpace(formatString))  fn.SetAnnotation("FormatString",  formatString);
    return fn.Name;
};

const string articleUrl = "https://www.daxpatterns.com/ranking/";
const string othersArticleUrl = "https://www.sqlbi.com/articles/showing-the-top-5-products-and-others-row/";
const string folderSuffixValue = "Ranking";
int createdCount = 0;

Func<string, string> EscapeDaxString = s => (s ?? "").Replace("\"", "\"\"");
Func<string, string> EscapeMString = s => (s ?? "").Replace("\"", "\"\"");
Func<Column, Table> CreateOthersRankingTable = col =>
{
    string tableName = $"{col.Table.Name} {col.Name} Rank";
    string rankingGroup = $"{rankingGroupPrefix} {col.Table.Name}";
    var table = Model.Tables.FirstOrDefault(t => t.Name == tableName);

    if (table != null)
    {
        var choice = MessageBox.Show(
            $"A table named '{tableName}' already exists.\n\n" +
            "Yes = Overwrite (delete and recreate the table and its relationships)\n" +
            "No = Reuse the existing table as-is\n" +
            "Cancel = Skip the Ranking with Others variant for this column",
            "Ranking with Others table already exists",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (choice == DialogResult.Cancel)
        {
            notes.Add($"Skipped 'Ranking with Others' for '{col.DaxObjectFullName}' because the existing table '{tableName}' was not replaced or reused.");
            return null;
        }

        if (choice == DialogResult.Yes)
        {
            table.Delete();
            table = null;
        }
    }

    if (table == null)
    {
        string columnRef = col.DaxObjectFullName;
        if (createOthersCalculatedTable)
        {
            string daxExpression = $@"UNION(
    SELECTCOLUMNS(
        ALLNOBLANKROW({columnRef}),
        ""Ranking name"", {columnRef},
        ""Ranking group"", ""{EscapeDaxString(rankingGroup)}""
    ),
    {{ ( ""{EscapeDaxString(othersName)}"", ""{EscapeDaxString(othersName)}"" ) }}
)";
            table = Model.AddCalculatedTable(tableName, daxExpression);
        }
        else
        {
            string sourceReference = $"#\"{EscapeMString(col.Table.Name)}\"";
            string sourceColumnName = EscapeMString(col.Name);
            string mExpression = $@"let
    Source = {sourceReference},
    #""Removed Other Columns"" = Table.SelectColumns(Source, {{""{sourceColumnName}""}}),
    #""Removed Duplicates"" = Table.Distinct(#""Removed Other Columns""),
    #""Filtered Rows"" = Table.SelectRows(#""Removed Duplicates"", each [{sourceColumnName}] <> null and [{sourceColumnName}] <> """"),
    #""Renamed Columns"" = Table.RenameColumns(#""Filtered Rows"", {{{{""{sourceColumnName}"", ""Ranking name""}}}}),
    #""Added Custom"" = Table.AddColumn(#""Renamed Columns"", ""Ranking group"", each ""{EscapeMString(rankingGroup)}"", type text),
    Others = #table(
        type table [Ranking name = text, Ranking group = text],
        {{{{""{EscapeMString(othersName)}"", ""{EscapeMString(othersName)}""}}}}
    ),
    #""Appended Query"" = Table.Combine({{#""Added Custom"", Others}})
in
    #""Appended Query""";
            // The false overload creates an import table, as used by the calendar macros.
            table = Model.AddTable(tableName, false);
            var partition = table.AddMPartition(tableName, mExpression);
            partition.Mode = ModeType.Import;
            table.AddDataColumn("Ranking name", "Ranking name", "", DataType.String);
            table.AddDataColumn("Ranking group", "Ranking group", "", DataType.String);
        }

        table.Description = $"Ranking table for {col.DaxObjectFullName}, including a '{othersName}' row.";
    }

    Column rankingNameColumn;
    Column rankingGroupColumn;
    if (!table.Columns.Contains("Ranking name") || !table.Columns.Contains("Ranking group"))
    {
        notes.Add($"Skipped 'Ranking with Others' UDF for '{col.DaxObjectFullName}' because table '{table.Name}' must contain 'Ranking name' and 'Ranking group' columns.");
        return null;
    }

    rankingNameColumn = table.Columns["Ranking name"];
    rankingGroupColumn = table.Columns["Ranking group"];
    rankingNameColumn.DataType = DataType.String;
    rankingNameColumn.SummarizeBy = AggregateFunction.None;
    rankingNameColumn.IsKey = true;
    rankingGroupColumn.DataType = DataType.String;
    rankingGroupColumn.SummarizeBy = AggregateFunction.None;

    var relationship = Model.Relationships.FirstOrDefault(r => r.FromColumn == col && r.ToColumn == rankingNameColumn);
    if (relationship == null)
    {
        relationship = Model.AddRelationship();
        relationship.FromColumn = col;
        relationship.ToColumn = rankingNameColumn;
        relationship.FromCardinality = RelationshipEndCardinality.Many;
        relationship.ToCardinality = RelationshipEndCardinality.One;
        relationship.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
        relationship.IsActive = true;
        notes.Add($"Created one-to-many relationship from {col.DaxObjectFullName} to {rankingNameColumn.DaxObjectFullName}.");
    }
    else
    {
        notes.Add($"Reused the existing relationship from {col.DaxObjectFullName} to {rankingNameColumn.DaxObjectFullName}.");
    }

    return table;
};

foreach (var col in pickedColumns)
{
    string columnRef = col.DaxObjectFullName;
    string tableRef  = col.Table.DaxObjectFullName;
    string colId     = SanitizeIdentifier(col.Name);
    if (string.IsNullOrEmpty(colId)) colId = "Column";
    Table othersRankingTable = wantOthersVariant ? CreateOthersRankingTable(col) : null;

    // -- 1) Rank position --------------------------------------------------------------------
    string rankFuncName = $"Local.Ranking.{colId}";
    string rankPrefix   = $"Ranking {col.Name} by";
    string rankBody = $@"(
    // {articleUrl}
    // What you are ranking - typically an amount/quantity measure to determine rank order
    baseMeasure: ANYREF
) =>
    IF(
        ISINSCOPE({columnRef}),
        VAR _CurrentValue = baseMeasure
        VAR _Rank =
            RANKX(
                ALLSELECTED({columnRef}),
                baseMeasure
            )
        RETURN
            IF(NOT ISBLANK(_CurrentValue), _Rank)
    )";
    string rankDescription = $"Dynamic rank position of baseMeasure within {columnRef} (blank outside the column's scope, or when baseMeasure is blank). Reference: {articleUrl}";

    string rankActualName = CreateLocalUDF(rankFuncName, rankBody, rankDescription, rankPrefix, "", folderSuffixValue, "#,0");
    if (rankActualName != null) createdCount++;

    // -- 2) Value if within Top N -------------------------------------------------------------
    if (rankActualName != null)
    {
        string topNFuncName = $"Local.Ranking.{colId}TopN";
        string topNSuffix   = $"by {col.Name}";
        string topNBody = $@"(
    // {articleUrl}
    // Base measure - used for ranking and returned as-is when within the Top N (e.g. profitability for the top N customers by sales), else blank
    baseMeasure: ANYREF
) =>
    VAR _Rank = {rankActualName}(baseMeasure)
    VAR _ValuesToRank = {topNArg}
    RETURN
        IF(
            _Rank > 0 && _Rank <= _ValuesToRank,
            baseMeasure
        )";
        string topNDescription = $"baseMeasure shown only when its rank (via {rankActualName}) is within the Top N ({topNArg}), else blank. Reference: {articleUrl}";

        if (CreateLocalUDF(topNFuncName, topNBody, topNDescription, topNPrefixLabel, topNSuffix, folderSuffixValue, "") != null) createdCount++;

        // -- 2b) Rank by one measure, summarize by another (2-parameter - not wrapped via the
        // "5 - Create measure from model dependent functions.cs" single-parameter macro, but
        // picked up automatically by "5 - Compare first selected measure with second selected
        // measure.cs" since it takes exactly two measure-compatible parameters).
        string byOtherFuncName = $"Local.Ranking.{colId}TopNbyOtherMeasure";
        string byOtherBody = $@"(
    // {articleUrl}
    // What to rank by, e.g. Sales Amount
    rankMeasure: ANYREF,
    // What to summarize/return for the rows within the Top N, e.g. Profit Margin
    summarizeMeasure: ANYREF
) =>
    VAR _Rank = {rankActualName}(rankMeasure)
    VAR _ValuesToRank = {topNArg}
    RETURN
        IF(
            _Rank > 0 && _Rank <= _ValuesToRank,
            summarizeMeasure
        )";
        string byOtherDescription = $"summarizeMeasure shown only for rows where rankMeasure's rank (via {rankActualName}) is within the Top N ({topNArg}), else blank - e.g. profit margin for the customers with the highest sales. Reference: {articleUrl}";

        if (CreateLocalUDF(byOtherFuncName, byOtherBody, byOtherDescription, "", "", folderSuffixValue, "") != null) createdCount++;
    }
    else
    {
        notes.Add($"Skipped the Top N and Top N-by-other-measure variants for column '{col.Name}' because its rank function was not created.");
    }

    // -- 2c) Top N with Others (optional) -----------------------------------------------------
    if (othersRankingTable != null)
    {
        string rankingNameRef = othersRankingTable.Columns["Ranking name"].DaxObjectFullName;
        string rankingGroupRef = othersRankingTable.Columns["Ranking group"].DaxObjectFullName;
        string othersFuncName = $"Local.Ranking.Rank_TopN_{colId}_with_others";
        string othersSuffix = $"by {col.Name} with {othersName}";
        string rankingGroup = $"{rankingGroupPrefix} {col.Table.Name}";
        string othersBody = $@"(
    // {othersArticleUrl}
    // The measure to aggregate for the Top N rows and the {EscapeDaxString(othersName)} remainder
    baseMeasure: ANYREF
) =>
    IF(
        ISINSCOPE({rankingGroupRef}),
        VAR _ValuesToRank = {topNArg}
        VAR _RankingGroup = SELECTEDVALUE({rankingGroupRef})
        VAR _TopRows =
            TOPN(
                _ValuesToRank,
                ALLSELECTED({othersRankingTable.DaxObjectFullName}),
                baseMeasure
            )
        RETURN
            SWITCH(
                _RankingGroup,
                ""{EscapeDaxString(rankingGroup)}"",
                    CALCULATE(baseMeasure, KEEPFILTERS(_TopRows)),
                ""{EscapeDaxString(othersName)}"",
                    IF(
                        NOT ISINSCOPE({rankingNameRef}),
                        VAR _TopAmount = CALCULATE(baseMeasure, _TopRows)
                        VAR _AllAmount =
                            CALCULATE(
                                baseMeasure,
                                ALLSELECTED({othersRankingTable.DaxObjectFullName})
                            )
                        RETURN
                            _AllAmount - _TopAmount
                    )
            ),
        baseMeasure
    )";
        string othersDescription = $"baseMeasure aggregated for the Top N {col.Name} values and the '{othersName}' remainder through {othersRankingTable.DaxObjectFullName}. Reference: {othersArticleUrl}";

        if (CreateLocalUDF(othersFuncName, othersBody, othersDescription, topNPrefixLabel, othersSuffix, folderSuffixValue, "") != null) createdCount++;
    }

    // -- 3) Ranking in Category (optional) ----------------------------------------------------
    if (wantCategoryVariant)
    {
        if (selectedCategoryColumn.Table != col.Table)
        {
            notes.Add($"Skipped the 'Ranking in Category' variant for column '{col.Name}' - the category column '{selectedCategoryColumn.Name}' is on a different table ('{selectedCategoryColumn.Table.Name}' vs '{col.Table.Name}').");
        }
        else
        {
            string categoryRef = selectedCategoryColumn.DaxObjectFullName;
            string catFuncName = $"Local.Ranking.{colId}InCategory";
            string catPrefix   = $"Ranking {col.Name} in {selectedCategoryColumn.Name} by";
            string catBody = $@"(
    // {articleUrl} - ""Rank in Category"" pattern
    // What you are ranking, restricted to rows sharing the same {categoryRef} value as the current row
    baseMeasure: ANYREF
) =>
    VAR _CurrentValue = baseMeasure
    VAR _ValuesToRank = {topNCategoryArg}
    VAR _CategoryScope =
        CALCULATETABLE(
            {tableRef},
            REMOVEFILTERS({columnRef}),
            ALLSELECTED({tableRef}),
            VALUES({categoryRef})
        )
    VAR _Rank =
        IF(
            ISINSCOPE({columnRef}),
            RANKX(_CategoryScope, baseMeasure)
        )
    RETURN
        IF(
            NOT ISBLANK(_CurrentValue) &&
            _Rank > 0 &&
            _Rank <= _ValuesToRank,
            _Rank
        )";
            string catDescription = $"Dynamic rank position of baseMeasure within {columnRef}, restricted to rows sharing the same {categoryRef} value and limited to the Top N ({topNCategoryArg}). Reference: {articleUrl}";

            if (CreateLocalUDF(catFuncName, catBody, catDescription, catPrefix, "", folderSuffixValue, "#,0") != null) createdCount++;
        }
    }
}

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = $"Done! Created/updated {createdCount} Ranking DaxUDF(s) for {pickedColumns.Count} column(s).\n\n" +
    "Next steps:\n" +
    "- Select a base measure and run '5 - Create measure from model dependent functions.cs' to wrap the Local.Ranking.<Column>, .<Column>TopN, .Rank_TopN_<Column>_with_others, and .<Column>InCategory functions into measures.\n" +
    "- Select two measures (rank by / summarize by) and run '5 - Compare first selected measure with second selected measure.cs' to wrap the Local.Ranking.<Column>TopNbyOtherMeasure function(s) into a measure.";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n => "- " + n));

Info(doneSummary);
