/*
 * Title: Create or update "Period to show" Calculation group (DaxUDF based)
 *
 * Author: Eivind Haugen
 *
 * Purpose: Dynamically switch the time period shown on an axis/visual using a disconnected calendar
 * table + calculation group.
 *
 * Reference: https://www.sqlbi.com/articles/show-previous-6-months-of-data-from-single-slicer-selection/
 * Modifications from the SQLBI article:
 *   1. Added KEEPFILTERS( _ShownDates ) to the window logic so it also works correctly with Calendar
 *      Options (not just the classic date table).
 *   2. Changed the in-window check to INTERSECT( VALUES(...), _ShownDates ) so it also works correctly
 *      with monthly/weekly/daily average measures.
 *
 * IMPORTANT: Review the calculation group's Precedence after running. It must be set HIGHER than any
 * Time Intelligence calculation group (so "Period to show" is applied on top of/after it), and
 * typically LOWER than e.g. a Unit Conversion calculation group (applied last, after the period/time
 * intelligence logic). Two calculation groups sharing the same Precedence give an undefined/unreliable
 * evaluation order - avoid ties. See the Info box shown the first time the calculation group is created.
 *
 * Flow (brief): Detects a disconnected calendar (a table filtering the main date table through an
 * inactive relationship: main Date = Many, disconnected Date = One; inactive fact-table relationships
 * are ignored) or creates one as a copy of the main date table, including the continuous axis settings
 * (data categories such as PaddedDateTableDates / Years, and the user hierarchies), cloning any Calendar
 * Options onto it so weekly/monthly logic resolves correctly; creates the required
 * model-specific ("Local.") DAX UDFs (reference date, year boundaries, period windows - reusing an
 * existing function/item instead of duplicating if an identical one already exists); and creates or
 * updates the "Period to show" calculation group and its calculation items based on the dialog
 * selections. Re-running against an already-linked calculation group/item skips the disconnected
 * calendar setup; running again with a different Calendar Option (Date source) adds a suffixed variant.
 *
 * NOTE: The week-based ("Previous N weeks") calculation items require a Calendar Option (e.g. a
 * "Weekly" ISO-week calendar) on the MAIN date table; a matching "<name> disconnected" clone of it is
 * created on the disconnected calendar (Step 2) so DATESINPERIOD/WEEK resolves correctly against the
 * disconnected calendar's own row context - verify in Power BI Desktop / TE3 after running, see the
 * Info messages at the end.
 */

#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

var warnings = new List<string>();

const string calcGroupLinkAnnotation = "PeriodToShow_DisconnectedTable";

// ============================================================================
// STEP 0 - Fast path: script run from an already-linked calculation group / item
// ============================================================================
// If a "Period to show" calculation group (or one of its calculation items) is already selected in
// the TOM Explorer, the disconnected-calendar setup (date column pick, relationship detect/create,
// Calendar Option cloning) only needs to happen once - skip it entirely and reuse what's already
// linked, jumping straight to picking which calculation items to add/update.
CalculationGroupTable calcGroupTable =
    Selected.Tables.OfType<CalculationGroupTable>().FirstOrDefault()
    ?? Selected.CalculationItems.Select(ci => ci.CalculationGroupTable).FirstOrDefault();

Column mainDateColumn = null;
Table  mainTable = null;
Table  disconnectedTable = null;
Column disconnectedDateColumn = null;
bool   disconnectedAlreadyExisted = false;

if (calcGroupTable != null)
{
    string linkedName = calcGroupTable.GetAnnotation(calcGroupLinkAnnotation);
    var linkedTable = !string.IsNullOrEmpty(linkedName) && Model.Tables.Contains(linkedName) ? Model.Tables[linkedName] : null;
    var linkedRel = linkedTable != null
        // The disconnected calendar filters the main date table: main date column (Many) -> disconnected (One)
        ? Model.Relationships.FirstOrDefault(r => !r.IsActive && r.ToTable == linkedTable
            && r.ToCardinality == RelationshipEndCardinality.One && r.FromColumn.DataType == DataType.DateTime)
        : null;

    if (linkedTable != null && linkedRel != null)
    {
        disconnectedTable = linkedTable;
        disconnectedDateColumn = linkedRel.ToColumn;
        mainDateColumn = linkedRel.FromColumn;
        mainTable = mainDateColumn.Table;
        disconnectedAlreadyExisted = true;
        Info($"Using the selected calculation group '{calcGroupTable.Name}', linked to disconnected calendar '{disconnectedTable.Name}' - relationship/calendar setup is skipped; only calculation items will be added/updated.");
    }
    else
    {
        warnings.Add($"The selected calculation group '{calcGroupTable.Name}' isn't linked to a disconnected calendar (or its relationship is missing) - falling back to the full setup flow.");
        calcGroupTable = null;
    }
}

bool needsFullSetup = mainDateColumn == null;

// Candidate date columns for the picker in the item-selection dialog below (only used if a full
// setup is required - i.e. no calculation group/item was already selected and linked).
var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime && c.IsKey).ToList();
if (needsFullSetup && !_dateColumns.Any())
{
    Error("No column marked as the key of a date table (IsKey = true) was found. Mark a date table first.");
    return;
}

// ============================================================================
// STEP 1 - Date source dialog (date column if needed, Classic vs one named Calendar Option)
// ============================================================================
var sourceForm = new Form { Text = "Select Date source", Width = 560, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
var sourceFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
sourceForm.Controls.Add(sourceFlow);

Action<string> AddSourceLabel = s => sourceFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });

// -- MAIN date table's date column - only needed the first time (a brand-new disconnected
// calendar/relationship has to be created); skipped entirely when re-running against an
// already-linked calculation group/item (see the FAST PATH in Step 0).
Button dateColButton = null;
if (needsFullSetup)
{
    AddSourceLabel("Date column of the MAIN (connected) date table:");
    dateColButton = new Button { Text = "Select Date column...", AutoSize = true };
    sourceFlow.Controls.Add(dateColButton);
}

// -- Date source: "Classic" (no Calendar Option, no name suffix) or ONE named Calendar Option on
// the main date table (suffixed by its name). Run the macro again with a different source selected
// if you need more than one variant (e.g. once for "Classic", once for "Weekly").
AddSourceLabel("Date source (run again with a different selection for additional variants):");
var sourceSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 420 };
sourceFlow.Controls.Add(sourceSelector);

Action RefreshSourceSelector = () => {
    sourceSelector.Items.Clear();
    sourceSelector.Items.Add("Classic");
    var cals = mainTable != null ? mainTable.Calendars.ToList() : new List<Calendar>();
    foreach (var cal in cals) sourceSelector.Items.Add(cal.Name);
    sourceSelector.SelectedIndex = 0;
};
RefreshSourceSelector();

if (needsFullSetup)
{
    dateColButton.Click += (s, e) => {
        var picked = SelectColumn(_dateColumns, null, "Select the Date-column from the MAIN date table");
        if (picked == null) return;
        mainDateColumn = picked;
        mainTable = picked.Table;
        dateColButton.Text = picked.DaxObjectFullName;
        RefreshSourceSelector();
    };
}

var sourceOk = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Margin = new Padding(0, 20, 8, 0) };
var sourceCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, Margin = new Padding(0, 20, 0, 0) };
var sourceBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
sourceBtnPanel.Controls.Add(sourceOk);
sourceBtnPanel.Controls.Add(sourceCancel);
sourceFlow.Controls.Add(sourceBtnPanel);
sourceForm.AcceptButton = sourceOk;
sourceForm.CancelButton = sourceCancel;

if (sourceForm.ShowDialog() != DialogResult.OK) { Info("Cancelled - no calculation items were created."); return; }

if (needsFullSetup && mainDateColumn == null)
{
    Error("No Date column was selected for the MAIN date table. Re-run and select one.");
    return;
}

Calendar selectedCalendar = mainTable.Calendars.FirstOrDefault(c => c.Name == (string)sourceSelector.SelectedItem);
bool isCalendarSource = selectedCalendar != null;
string nameSuffix = isCalendarSource ? " - " + selectedCalendar.Name : "";
Func<string, string> Sanitize = s => new string(s.Where(char.IsLetterOrDigit).ToArray());
string fnSuffix = isCalendarSource ? "_" + Sanitize(selectedCalendar.Name) : "";
bool sourceHasWeek = isCalendarSource && selectedCalendar.GetTimeUnits().Any(tu => tu.TimeUnit == TimeUnit.Week);
bool sourceHasMonth = !isCalendarSource || selectedCalendar.GetTimeUnits().Any(tu => tu.TimeUnit == TimeUnit.Month || tu.TimeUnit == TimeUnit.MonthOfYear);

// Auto-detect the Year column from the selected Calendar Option's Year time unit; falls back to a
// manual picker in the details dialog (Step 1b) below if not found, or if "Classic" was selected.
Column yearColumn = isCalendarSource
    ? (selectedCalendar.GetTimeUnits().FirstOrDefault(tu => tu.TimeUnit == TimeUnit.Year)?.PrimaryColumn)
    : null;

// ============================================================================
// STEP 1a - Item selection dialog (which calculation items to include)
// ============================================================================
var itemForm = new Form { Text = "Select 'Period to show' calculation items" + nameSuffix, Width = 560, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
var flowPanel = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
itemForm.Controls.Add(flowPanel);

Action<string> AddLabel = s => flowPanel.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });
AddLabel("Items to include (N-values, if any, are collected in the next dialog):");

var ytdCheck        = new CheckBox { Text = "This year to selected date",                      AutoSize = true, Checked = false };
var entireYearCheck = new CheckBox { Text = "Show entire selected year",                        AutoSize = true, Checked = false };
var upToTodayCheck  = new CheckBox { Text = "Show up to today and last year (year-end anchor)", AutoSize = true, Checked = false };
var remainingCheck  = new CheckBox { Text = "Remaining year",                                   AutoSize = true, Checked = false };
var yearsBackCheck  = new CheckBox { Text = "Year-to-date + last N years",                       AutoSize = true, Checked = false };
var nextMonthsCheck = new CheckBox { Text = "Next N months",                                    AutoSize = true, Checked = false, Enabled = sourceHasMonth };
var monthsCheck     = new CheckBox { Text = "Previous N months",                                AutoSize = true, Checked = false, Enabled = sourceHasMonth };
var weeksCheck      = new CheckBox { Text = "Previous N weeks",                                 AutoSize = true, Checked = false, Enabled = sourceHasWeek };
var nextWeeksCheck  = new CheckBox { Text = "Next N weeks",                                     AutoSize = true, Checked = false, Enabled = sourceHasWeek };
foreach (var c in new[] { ytdCheck, entireYearCheck, upToTodayCheck, remainingCheck, yearsBackCheck, nextMonthsCheck, monthsCheck, weeksCheck, nextWeeksCheck }) flowPanel.Controls.Add(c);
if (!sourceHasMonth)
    flowPanel.Controls.Add(new Label { Margin = new Padding(20, 0, 0, 5), AutoSize = true,
        Text = "(Previous/Next N months needs a Calendar Option with a Month time unit as the Date source above.)" });
if (!sourceHasWeek)
    flowPanel.Controls.Add(new Label { Margin = new Padding(20, 0, 0, 5), AutoSize = true,
        Text = "(Previous/Next N weeks needs a Calendar Option with a Week time unit as the Date source above.)" });

var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Margin = new Padding(0, 20, 8, 0) };
var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, Margin = new Padding(0, 20, 0, 0) };
var buttonPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
buttonPanel.Controls.Add(okButton);
buttonPanel.Controls.Add(cancelButton);
flowPanel.Controls.Add(buttonPanel);
itemForm.AcceptButton = okButton;
itemForm.CancelButton = cancelButton;

if (itemForm.ShowDialog() != DialogResult.OK) { Info("Cancelled - no calculation items were created."); return; }

bool yearItemsChecked = ytdCheck.Checked || entireYearCheck.Checked || remainingCheck.Checked
    || upToTodayCheck.Checked || yearsBackCheck.Checked;

Func<string, List<int>> ParseNList = text => text.Split(',')
    .Select(s => s.Trim())
    .Where(s => int.TryParse(s, out _))
    .Select(int.Parse)
    .Distinct()
    .ToList();

// ============================================================================
// STEP 1b - Details dialog: N-value inputs and any manual Year-column fallback
// ============================================================================
var detailForm = new Form { Text = "'Period to show' details" + nameSuffix, Width = 560, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
var detailFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
detailForm.Controls.Add(detailFlow);

Action<string> AddDetailLabel = s => detailFlow.Controls.Add(new Label { Margin = new Padding(0, 10, 0, 5), Text = s, AutoSize = true });
Func<string, string, TextBox> AddDetailNRow = (label, defaultN) => {
    var panel = new Panel { AutoSize = true };
    panel.Controls.Add(new Label { Text = label, AutoSize = true, Location = new Point(0, 4) });
    var box = new TextBox { Location = new Point(280, 0), Width = 120, Text = defaultN };
    panel.Controls.Add(box);
    detailFlow.Controls.Add(panel);
    return box;
};

bool needsYearColFallback = yearItemsChecked && yearColumn == null;
bool anyDetailControls = (monthsCheck.Checked && sourceHasMonth) || (nextMonthsCheck.Checked && sourceHasMonth) || yearsBackCheck.Checked
    || (weeksCheck.Checked && sourceHasWeek) || (nextWeeksCheck.Checked && sourceHasWeek) || needsYearColFallback;

TextBox monthsBox = null, nextMonthsBox = null, yearsBackBox = null, weeksBox = null, nextWeeksBox = null;

if (anyDetailControls)
{
    if ((monthsCheck.Checked && sourceHasMonth) || (nextMonthsCheck.Checked && sourceHasMonth) || yearsBackCheck.Checked
        || (weeksCheck.Checked && sourceHasWeek) || (nextWeeksCheck.Checked && sourceHasWeek))
        AddDetailLabel("N-value lists accept one or more comma-separated numbers (e.g. \"3,6,12\"):");
    if (monthsCheck.Checked && sourceHasMonth)     monthsBox     = AddDetailNRow("Previous N months",         "3,6,12");
    if (nextMonthsCheck.Checked && sourceHasMonth) nextMonthsBox = AddDetailNRow("Next N months",             "6");
    if (yearsBackCheck.Checked)  yearsBackBox  = AddDetailNRow("Year-to-date + last N years", "1");
    if (weeksCheck.Checked && sourceHasWeek) weeksBox = AddDetailNRow("Previous N weeks",    "12");
    if (nextWeeksCheck.Checked && sourceHasWeek) nextWeeksBox = AddDetailNRow("Next N weeks", "6");

    if (needsYearColFallback)
    {
        AddDetailLabel("Year column (used for 'This year' / 'Entire year' / 'Remaining year' / 'Year-to-date + last N years'):");
        var yearColButton = new Button { Text = "Select Year column...", AutoSize = true };
        yearColButton.Click += (s, e) => {
            var yearCols = mainTable.Columns.Where(c => c.Name.IndexOf("year", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            if (!yearCols.Any()) yearCols = mainTable.Columns.ToList();
            var picked = SelectColumn(yearCols, null, "Select the Year column from the main date table");
            if (picked != null) { yearColumn = picked; yearColButton.Text = picked.DaxObjectFullName; }
        };
        detailFlow.Controls.Add(yearColButton);
    }

    var detailOk = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Margin = new Padding(0, 20, 8, 0) };
    var detailCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, Margin = new Padding(0, 20, 0, 0) };
    var detailButtonPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
    detailButtonPanel.Controls.Add(detailOk);
    detailButtonPanel.Controls.Add(detailCancel);
    detailFlow.Controls.Add(detailButtonPanel);
    detailForm.AcceptButton = detailOk;
    detailForm.CancelButton = detailCancel;

    if (detailForm.ShowDialog() != DialogResult.OK) { Info("Cancelled - no calculation items were created."); return; }
}

if (yearItemsChecked && yearColumn == null)
{
    warnings.Add($"No Year column was selected - 'This year' / 'Entire year' / 'Remaining year' / 'Year-to-date + last N years' items{nameSuffix} were skipped.");
}

if (needsFullSetup)
{
// ============================================================================
// STEP 2 - Detect (or create) the disconnected calendar (only on a full/first-time setup run)
// ============================================================================
// Only a table that FILTERS the main date table counts: an inactive relationship where the main date
// column is the Many side (From) and a date column on another table is the One side (To). Inactive
// relationships from fact tables to the date table (fact = Many, date = One) are ignored, and existing
// relationships are never modified.
var existingInactiveRel = Model.Relationships.FirstOrDefault(r =>
    !r.IsActive
    && r.FromColumn == mainDateColumn
    && r.ToCardinality == RelationshipEndCardinality.One
    && r.ToTable != mainTable
    && r.ToColumn.DataType == DataType.DateTime);

disconnectedAlreadyExisted = existingInactiveRel != null;

if (existingInactiveRel != null)
{
    disconnectedDateColumn = existingInactiveRel.ToColumn;
    disconnectedTable = disconnectedDateColumn.Table;
    Info($"Found existing disconnected calendar: '{disconnectedTable.Name}' (inactive relationship {mainDateColumn.DaxObjectFullName} -> {disconnectedDateColumn.DaxObjectFullName}).");
}
else
{
    string newTableName = Interaction.InputBox(
        "No disconnected calendar (inactive relationship) was found for this date table." + Environment.NewLine +
        "Provide a name for the new disconnected calendar table:",
        "Create Disconnected Calendar", mainTable.Name + " disconnected");

    if (string.IsNullOrWhiteSpace(newTableName)) { Info("No table name provided. Operation cancelled."); return; }
    if (Model.Tables.Contains(newTableName))
    {
        Error($"A table named '{newTableName}' already exists. Re-run and pick a different name, or delete/rename it first.");
        return;
    }

    Table newTable = null;
    bool isMPartitionTable = false;
    try
    {
        string mExpr = $@"let
    Source = #""{mainTable.Name}""
in
    Source";
        newTable = Model.AddTable(newTableName, false);
        var partition = newTable.AddMPartition(newTableName, mExpr);
        partition.Mode = ModeType.Import;
        isMPartitionTable = true;
    }
    catch (Exception ex)
    {
        warnings.Add($"Write-operation not supported ({ex.Message}); created a calculated table instead of an M-partition table.");
        newTable = Model.AddCalculatedTable(newTableName, mainTable.DaxObjectFullName);
    }

    newTable.Description = $"Disconnected copy of '{mainTable.Name}', used for the 'Period to show' calculation group (not connected via an active relationship).";

    // ---- For an M-partition table, TOM does not auto-detect the schema from the M query - each
    // column must be explicitly declared (name/source-column must match what the M query returns),
    // same pattern as used in '1 - Create English Calendar (combined).cs'. Calculated tables infer
    // their columns automatically once evaluated, so this step is skipped for that fallback path.
    if (isMPartitionTable)
    {
        foreach (var srcCol in mainTable.Columns)
        {
            if (newTable.Columns.Contains(srcCol.Name)) continue;
            newTable.AddDataColumn(srcCol.Name, srcCol.Name, srcCol.DisplayFolder, srcCol.DataType);
        }
    }

    // ---- Copy column-level metadata (description, folder, summarize-by, format string, hidden) ----
    // Data categories (PaddedDateTableDates / Years) and hierarchies are copied further below.
    foreach (var srcCol in mainTable.Columns)
    {
        if (!newTable.Columns.Contains(srcCol.Name)) continue;
        var dstCol = newTable.Columns[srcCol.Name];
        dstCol.Description   = srcCol.Description;
        dstCol.DisplayFolder = srcCol.DisplayFolder;
        dstCol.SummarizeBy   = srcCol.SummarizeBy;
        dstCol.FormatString  = srcCol.FormatString;
        dstCol.IsHidden      = srcCol.IsHidden;
    }
    // Best practice: short date format for any date column that still has no explicit format string
    // (e.g. defaulting to "Long Date"), same convention as "1 - Create English Calendar (combined).cs".
    foreach (var dstCol in newTable.Columns)
    {
        if (dstCol.DataType == DataType.DateTime && string.IsNullOrEmpty(dstCol.FormatString))
            dstCol.FormatString = "MM/dd/yyyy";
    }
    // Second pass for SortByColumn, so the referenced column already exists on the new table
    foreach (var srcCol in mainTable.Columns)
    {
        var srcDataCol = srcCol as DataColumn;
        if (srcDataCol == null || srcDataCol.SortByColumn == null) continue;
        if (!newTable.Columns.Contains(srcCol.Name) || !newTable.Columns.Contains(srcDataCol.SortByColumn.Name)) continue;

        (newTable.Columns[srcCol.Name] as DataColumn).SortByColumn =
            newTable.Columns[srcDataCol.SortByColumn.Name] as DataColumn;
    }

    // Note: Calendar Options ("calendar" TOM objects, e.g. Gregorian/Weekly) are cloned onto this
    // disconnected table further below (once the inactive relationship is in place), named
    // "<original name> disconnected" - required so the weekly calculation items can use
    // DATESINPERIOD/WEEK correctly against the disconnected calendar's own row context.

    // ---- Pick the corresponding date column & add the inactive relationship ----
    if (!newTable.Columns.Contains(mainDateColumn.Name))
    {
        Error($"Could not find a column named '{mainDateColumn.Name}' on the new table '{newTable.Name}'.");
        return;
    }
    disconnectedTable = newTable;
    disconnectedDateColumn = newTable.Columns[mainDateColumn.Name];

    // ---- Mark as a date table, same as the main date table (DataCategory = "Time" + date key column) ----
    disconnectedTable.DataCategory = "Time";
    disconnectedDateColumn.IsKey = true;
    disconnectedDateColumn.DataType = DataType.DateTime;

    try
    {
        var rel = Model.AddRelationship();
        rel.FromColumn = mainDateColumn;
        rel.ToColumn   = disconnectedDateColumn;
        // From: main date table (Many), To: disconnected calendar (One), single direction, once
        // activated via USERELATIONSHIP in the DAX below.
        rel.FromCardinality = RelationshipEndCardinality.Many;
        rel.ToCardinality   = RelationshipEndCardinality.One;
        rel.IsActive = false;
    }
    catch (Exception ex)
    {
        Error($"Disconnected table created, but failed to add the inactive relationship: {ex.Message}");
        return;
    }
}

// ---- Continuous axis settings (SQLBI: continuous axis for any date table), copied from the main date
// table: column data categories (e.g. PaddedDateTableDates on Date, Years on Year) and user hierarchies
// (e.g. Year-Quarter-Month), matched by column name. Only fills gaps, so it is safe on an existing
// disconnected calendar: categories already set and hierarchies with the same name are left as they are.
foreach (var srcCol in mainTable.Columns)
{
    if (string.IsNullOrEmpty(srcCol.DataCategory) || !disconnectedTable.Columns.Contains(srcCol.Name)) continue;
    var dstCol = disconnectedTable.Columns[srcCol.Name];
    if (string.IsNullOrEmpty(dstCol.DataCategory)) dstCol.DataCategory = srcCol.DataCategory;
}
foreach (var srcHier in mainTable.Hierarchies)
{
    if (disconnectedTable.Hierarchies.Contains(srcHier.Name)) continue;
    var srcLevels = srcHier.Levels.OrderBy(l => l.Ordinal).ToList();
    if (srcLevels.Any(l => !disconnectedTable.Columns.Contains(l.Column.Name)))
    {
        warnings.Add($"Hierarchy '{srcHier.Name}' was not copied to '{disconnectedTable.Name}': not all of its level columns exist there.");
        continue;
    }
    // Empty Column[] picks the Column overload (without it the call is ambiguous with the string[] one)
    var dstHier = disconnectedTable.AddHierarchy(srcHier.Name, srcHier.DisplayFolder, new Column[0]);
    dstHier.Description = srcHier.Description;
    dstHier.IsHidden    = srcHier.IsHidden;
    foreach (var lvl in srcLevels)
        dstHier.AddLevel(disconnectedTable.Columns[lvl.Column.Name], lvl.Name);
}

// ---- Ensure a matching "<name> disconnected" Calendar Option exists on the disconnected calendar for
// every Calendar Option (e.g. Weekly, Gregorian) on the main date table. The weekly calculation items
// evaluate MAX()/DATESINPERIOD() in the row context of the DISCONNECTED calendar, so DATESINPERIOD
// needs a Calendar Option defined on that same table - reusing the main table's Calendar Option here
// causes "a single value for column [Date] ... cannot be determined" errors. Idempotent: skipped if a
// matching Calendar Option already exists (e.g. on a re-run).
foreach (var srcCal in mainTable.Calendars)
{
    string disconnectedCalName = srcCal.Name + " disconnected";
    if (disconnectedTable.Calendars.Any(c => c.Name.Equals(disconnectedCalName, StringComparison.OrdinalIgnoreCase))) continue;

    // Every column the Calendar Option uses must exist (by name) on the disconnected table; otherwise
    // skip it rather than create an incomplete copy.
    var srcTimeUnits = srcCal.GetTimeUnits().ToList();
    var missingCalCols = srcTimeUnits
        .SelectMany(tu => new[] { tu.PrimaryColumn }.Concat(tu.AssociatedColumns))
        .Where(c => c != null && !disconnectedTable.Columns.Contains(c.Name))
        .Select(c => c.Name).Distinct().ToList();
    if (missingCalCols.Any())
    {
        warnings.Add($"Calendar Option '{srcCal.Name}' was not copied to '{disconnectedTable.Name}': missing column(s) {string.Join(", ", missingCalCols)}.");
        continue;
    }
    try
    {
        // Rebuilt from scratch on the disconnected table (same calls as the date table macro), never with
        // Clone(): a clone keeps references to the MAIN table's columns, and clearing its associated
        // columns to re-point them deletes those columns from the main date table.
        var newCal = disconnectedTable.AddCalendar(disconnectedCalName);
        newCal.Description = srcCal.Description;
        foreach (var tu in srcTimeUnits.Where(t => t.PrimaryColumn != null))
        {
            var associated = tu.AssociatedColumns.Select(c => disconnectedTable.Columns[c.Name]).ToArray();
            newCal.AddTimeUnit(tu.TimeUnit, disconnectedTable.Columns[tu.PrimaryColumn.Name], associated);
        }
    }
    catch (Exception ex)
    {
        warnings.Add($"Could not copy Calendar Option '{srcCal.Name}' to '{disconnectedTable.Name}' as '{disconnectedCalName}': {ex.Message}. Weekly calculation items relying on it may fail - consider copying the Calendar Option to the disconnected calendar manually in TE3.");
    }
}

// Group the disconnected calendar table together with other dimension tables (matches the source
// date table's grouping convention).
disconnectedTable.SetAnnotation("TabularEditor_TableGroup", "Dimension");
} // end of "if (needsFullSetup)" full setup flow

string mainDateCol       = mainDateColumn.DaxObjectFullName;       // e.g. 'Calendar'[Date]
string mainTableRef      = mainTable.DaxObjectFullName;            // e.g. 'Calendar'
string disconnectedDateCol  = disconnectedDateColumn.DaxObjectFullName;
string disconnectedTableRef = disconnectedTable.DaxObjectFullName;

// ============================================================================
// STEP 3 - Calculation group: detect / append / create (skipped if resolved via the fast path above)
// ============================================================================
const string calcGroupDefaultName = "Period to show";
bool isNewCalcGroupTable = false;

if (calcGroupTable == null && disconnectedAlreadyExisted)
{
    // Only worth searching for / prompting about an existing calculation group if the disconnected
    // calendar itself already existed - a brand-new disconnected calendar can't have one linked yet.

    // Auto-detect a previously-created calculation group that is already linked (via annotation) to
    // this disconnected calendar, so re-running the script updates the right group even if it (or its
    // calculation items) were renamed in the meantime.
    var linkedCgTable = Model.CalculationGroups.FirstOrDefault(cg =>
        cg.GetAnnotation(calcGroupLinkAnnotation) == disconnectedTable.Name);

    if (linkedCgTable != null)
    {
        calcGroupTable = linkedCgTable;
        Info($"Found existing '{calcGroupTable.Name}' calculation group already linked to '{disconnectedTable.Name}' - its calculation items will be updated/appended.");
    }
    else
    {
        var existingCgTables = Model.CalculationGroups.ToList();
        if (existingCgTables.Any())
        {
            var response = MessageBox.Show(
                $"Found {existingCgTables.Count} existing calculation group(s) in the model, none of which are linked to '{disconnectedTable.Name}'." + Environment.NewLine +
                "Click 'Yes' to update/append calculation items to one of the existing calculation groups, or 'No' to create a new one.",
                "Calculation group", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (response == DialogResult.Yes)
            {
                var picked = SelectTable(existingCgTables.Cast<Table>().ToList(), null, "Select the existing calculation group to update/append items in");
                calcGroupTable = picked as CalculationGroupTable;
                if (calcGroupTable == null) Info("No existing calculation group selected - a new one will be created instead.");
            }
        }
    }
}

if (calcGroupTable == null)
{
    isNewCalcGroupTable = true;
    string tableName = calcGroupDefaultName;
    int suffix = 1;
    while (Model.Tables.Contains(tableName)) { tableName = calcGroupDefaultName + " " + (++suffix); }

    calcGroupTable = Model.AddCalculationGroupTable(tableName);
    (calcGroupTable.Columns["Name"] as DataColumn).Name = "Time Frame";
    (calcGroupTable.Columns["Time Frame"] as DataColumn).Description = "The time frame to show on axis when the disconnected calendar is used";
    calcGroupTable.Description = "Calculation group to dynamically change the time frame shown on axis, using a disconnected calendar (DAX UDF based)";
    calcGroupTable.CalculationGroup.Precedence = 1;
}

// Link this calculation group to the disconnected calendar so future runs can auto-detect it, even if
// the table/items get renamed afterwards.
calcGroupTable.SetAnnotation(calcGroupLinkAnnotation, disconnectedTable.Name);

int nextOrdinal = calcGroupTable.CalculationItems.Any() ? calcGroupTable.CalculationItems.Max(ci => ci.Ordinal) + 1 : 0;

Action<string, string, string> AddCalcItem = (name, expression, description) =>
{
    // If a differently-named item already has the exact same expression, skip creating a visual duplicate.
    var identicalElsewhere = calcGroupTable.CalculationItems.FirstOrDefault(ci =>
        !ci.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && ci.Expression == expression);
    if (identicalElsewhere != null)
    {
        warnings.Add($"Skipped creating '{name}' - calculation item '{identicalElsewhere.Name}' already has an identical expression.");
        return;
    }

    var existingItem = calcGroupTable.CalculationItems.FirstOrDefault(ci => ci.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    if (existingItem != null)
    {
        if (existingItem.Expression == expression) return; // already up to date - nothing to do, no prompt

        var response = MessageBox.Show(
            $"Calculation item '{name}' already exists in '{calcGroupTable.Name}' with a different expression." + Environment.NewLine + Environment.NewLine +
            "Click 'Yes' to overwrite it, or 'No' to skip and leave it unchanged.",
            "Calculation item already exists", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (response == DialogResult.Yes)
        {
            existingItem.Expression  = expression;
            existingItem.Description = description;
            warnings.Add($"Calculation item '{name}' already existed in '{calcGroupTable.Name}' with a different expression - overwritten.");
        }
        else
        {
            warnings.Add($"Calculation item '{name}' already existed in '{calcGroupTable.Name}' with a different expression - left unchanged (skipped).");
        }
        return;
    }
    var item = calcGroupTable.AddCalculationItem(name);
    item.Expression  = expression;
    item.Ordinal     = nextOrdinal++;
    item.Description = description;
};

// ============================================================================
// STEP 4 - Create / reuse the required DAX UDFs
// ============================================================================
const string localPrefix = "Local";

// Fully macro-owned ("Local.") functions are safe to recreate so they always match current selections.
// Before creating/renaming, also check whether a DIFFERENTLY-NAMED function already has the exact same
// DAX expression (e.g. two Calendar Options that happen to resolve to the same Year column produce a
// byte-identical FirstDayOfYear body) - if so, reuse that existing function instead of creating a
// duplicate. Returns the name that should actually be used by the caller (may differ from "name").
Func<string, string, string, string> EnsureLocalFunction = (name, body, description) =>
{
    var identicalElsewhere = Model.Functions.FirstOrDefault(f => f.Name != name && f.Expression == body);
    if (identicalElsewhere != null)
    {
        warnings.Add($"'{name}' has an identical DAX expression to existing function '{identicalElsewhere.Name}' - reusing it instead of creating a duplicate.");
        return identicalElsewhere.Name;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == name);
    if (existing != null)
    {
        if (existing.Expression == body) return name; // already up to date - nothing to do
        existing.Delete();
    }
    var fn = Model.AddFunction(name);
    fn.Expression  = body;
    fn.Description = description;
    return name;
};

string refDateFuncName = localPrefix + ".DateControlVisual.ReferenceInput.LastHistoricalDate";
refDateFuncName = EnsureLocalFunction(refDateFuncName, $@"(
    // No input required
) =>
    MIN( MAX( {mainDateCol} ), TODAY() )",
    "Model specific reference date - the last selected date in the main date table, capped at today");

// No generic/reusable layer - these are created directly as model-specific ("Local.") functions,
// since a shared cross-model wrapper added complexity without real value here. The year column is
// baked directly into the function body (rather than passed in by each calculation item), and a
// yearOffset argument (0 = the selected/current year, -1 = last year, -N = N years back) generalizes
// "this year" / "last year" / "N years back" through a single function pair, correctly reflecting the
// selected Date source's own year boundaries (important for fiscal/ISO calendars where a year doesn't
// necessarily run Jan-Dec) rather than assuming a hardcoded Gregorian DATE(YEAR()+/-n, ...) shift.
Func<string, string, string> BuildYearBoundaryBody = (aggFunc, yearColRef) => $@"(
    // 0 for the selected/current year, -1 for last year, -N for N years back, 1 for next year, etc.
    // A fixed INT64 is used here since this is a hard-coded offset per calculation item (not driven
    // by a measure) - change back to ""ANYREF EXPR"" if you need to pass a dynamic value computed from
    // a measure (e.g. a slicer-driven offset).
    yearOffset : INT64
) =>
    VAR _ReferenceDate = {refDateFuncName}()
    VAR _CurrentYearValue =
        CALCULATE(
            MAX( {yearColRef} ),
            REMOVEFILTERS( {mainTableRef} ),
            {mainDateCol} = _ReferenceDate
        )
    VAR _TargetYear = _CurrentYearValue + yearOffset
    VAR _{aggFunc}Date =
        CALCULATE(
            {aggFunc}( {mainDateCol} ),
            REMOVEFILTERS( {mainTableRef} ),
            FILTER(
                ALL( {yearColRef} ),
                {yearColRef} = _TargetYear
            )
        )
    RETURN
        _{aggFunc}Date";

// FirstDayOfYear/LastDayOfYear for the selected Date source - no suffix for "Classic", suffixed by
// Calendar Option name otherwise (e.g. FirstDayOfYear_Weekly), since a different Date source can have
// a different "last selected date" and therefore different year boundaries.
string firstDayFuncName = null, lastDayFuncName = null;
if (yearColumn != null)
{
    string yearColRef = yearColumn.DaxObjectFullName;
    firstDayFuncName = localPrefix + ".DateControlVisual.ReferenceInput.FirstDayOfYear" + fnSuffix;
    lastDayFuncName  = localPrefix + ".DateControlVisual.ReferenceInput.LastDayOfYear" + fnSuffix;
    firstDayFuncName = EnsureLocalFunction(firstDayFuncName, BuildYearBoundaryBody("MIN", yearColRef),
        $"Model specific: first date of the year ({yearColRef}), offset from the last selected date's year{nameSuffix}");
    lastDayFuncName  = EnsureLocalFunction(lastDayFuncName, BuildYearBoundaryBody("MAX", yearColRef),
        $"Model specific: last date of the year ({yearColRef}), offset from the last selected date's year{nameSuffix}");
}

// Builds the "start/end window" DAX UDF body (blank outside the selected window). Kept as a shared
// "StartEnd" DAX UDF (baseMeasure/startDateDefinition/endDateDefinition) - each calculation item just
// calls it with its own start/end date definitions, so it's easy to update the window logic in a
// single place if required.
string startEndFunc = localPrefix + ".DateControlVisual.PeriodToShow.StartEnd";
startEndFunc = EnsureLocalFunction(startEndFunc, $@"(
    // Common use is in a calculation group, using SELECTEDMEASURE(). Kept as ANYREF (not MEASUREREF)
    // because SELECTEDMEASURE() is a function call, not a literal measure reference, so MEASUREREF
    // would reject the calculation-group usage - only switch to MEASUREREF if this is only ever
    // called directly with an existing measure (e.g. [Sales Amount]), never with SELECTEDMEASURE().
    baseMeasure : ANYREF,
    // The first date to include in the window, e.g. the first date of the selected year
    startDateDefinition : ANYREF,
    // The last date to include in the window, e.g. the last selected date
    endDateDefinition : ANYREF
) =>
    // Reference article: https://www.sqlbi.com/articles/show-previous-6-months-of-data-from-single-slicer-selection/
    VAR _StartDate = startDateDefinition
    VAR _EndDate   = endDateDefinition
    VAR _ShownDates =
        DATESBETWEEN(
            {disconnectedDateCol},
            _StartDate,
            _EndDate
        )
    VAR _IsInWindow =
        COUNTROWS(
            INTERSECT( VALUES( {disconnectedDateCol} ), _ShownDates )
        ) > 0
    VAR _Result =
        IF(
            _IsInWindow,
            CALCULATE(
                baseMeasure,
                CALCULATETABLE(
                    VALUES( {mainDateCol} ),
                    REMOVEFILTERS( {mainTableRef} ),
                    KEEPFILTERS( _ShownDates ),
                    USERELATIONSHIP( {mainDateCol}, {disconnectedDateCol} )
                )
            )
        )
    RETURN
        _Result",
    "Shows a measure only within the given start/end date window (blank outside it). Modified from the SQLBI article: added KEEPFILTERS so it also works with Calendar Options, and checks the window with INTERSECT so it also works with monthly/weekly/daily average measures. Works with both classic and Calendar Option date sources.");

// Builds the calculation item expression that calls the shared StartEnd DAX UDF with the given
// start/end date DAX expressions, substituting SELECTEDMEASURE() as the base measure.
Func<string, string, string> BuildStartEndBody = (startExpr, endExpr) => $@"{startEndFunc}(
    SELECTEDMEASURE(),
    {startExpr},
    {endExpr}
)";

// Quotes a name as a DAX identifier (table/calendar reference syntax), doubling any embedded quotes.
Func<string, string> QuoteDaxIdentifier = name => "'" + name.Replace("'", "''") + "'";

// When a Calendar Option is the selected Date source, DATESINPERIOD's WEEK interval requires a
// *calendar reference* (not a column reference) as its first argument - a column reference only
// accepts DAY/MONTH/QUARTER/YEAR. Reusing the same calendar-reference source for MONTH too gives
// consistent behavior across both and matches how Microsoft's calendar-based time intelligence is
// designed to be used. This also requires the "Enhanced DAX Time Intelligence" preview feature to be
// enabled in Power BI Desktop (File > Options and settings > Options > Preview features) - otherwise
// DATESINPERIOD rejects WEEK entirely regardless of source syntax.
// For "Classic" (no Calendar Option selected), there's no calendar to reference, so MONTH keeps using
// the classic column-based source (WEEK is unavailable in that case - sourceHasWeek is false).
string disconnectedCalendarRef = isCalendarSource
    ? QuoteDaxIdentifier(selectedCalendar.Name + " disconnected")
    : null;

// Builds the DAX body for a rolling N-period window (blank outside the window), for either MONTH or
// WEEK granularity. Uses the disconnected calendar's own Calendar Option (a calendar reference, e.g.
// 'Weekly disconnected') as the DATESINPERIOD source when a Calendar Option is selected - required for
// WEEK, and used for MONTH too for consistency - or the disconnected Date column for "Classic". The
// in-window check follows suit: CALCULATETABLE(disconnected table, KEEPFILTERS(_ShownDates)) for a
// calendar reference (SELECTCOLUMNS/most table functions aren't supported on calendar-based time
// intelligence results), or the original INTERSECT(VALUES(...), ...) for the classic column source.
Func<string, string, string> BuildPeriodFunction = (unitLabel, unitKeyword) =>
{
    string datesInPeriodSource = isCalendarSource ? disconnectedCalendarRef : disconnectedDateCol;
    string isInWindowBody = isCalendarSource
        ? $@"COUNTROWS(
            CALCULATETABLE( {disconnectedTableRef}, KEEPFILTERS( _ShownDates ) )
        ) > 0"
        : $@"COUNTROWS(
            INTERSECT( VALUES( {disconnectedDateCol} ), _ShownDates )
        ) > 0";

    return $@"(
    // Common use is in a calculation group, using SELECTEDMEASURE(). Kept as ANYREF (not MEASUREREF)
    // because SELECTEDMEASURE() is a function call, not a literal measure reference, so MEASUREREF
    // would reject the calculation-group usage - only switch to MEASUREREF if this is only ever
    // called directly with an existing measure (e.g. [Sales Amount]), never with SELECTEDMEASURE().
    baseMeasure : ANYREF,
    // Positive to look forward, negative to look backward, e.g. -6 for the previous 6 {unitLabel.ToLower()}.
    // A fixed INT64 is used here since this is a hard-coded value per calculation item (not driven by
    // a measure) - change back to ""ANYREF EXPR"" if you need to pass a dynamic value computed from a
    // measure (e.g. a slicer-driven number of {unitLabel.ToLower()}).
    numberOf{unitLabel} : INT64
) =>
    // Reference article: https://www.sqlbi.com/articles/show-previous-6-months-of-data-from-single-slicer-selection/
    VAR _Number    = numberOf{unitLabel}
    VAR _Reference = {refDateFuncName}()
    VAR _ShownDates =
        DATESINPERIOD(
            {datesInPeriodSource},
            _Reference,
            _Number,
            {unitKeyword}
        )
    VAR _IsInWindow =
        {isInWindowBody}
    VAR _Result =
        IF(
            _IsInWindow,
            CALCULATE(
                baseMeasure,
                CALCULATETABLE(
                    VALUES( {mainDateCol} ),
                    REMOVEFILTERS( {mainTableRef} ),
                    KEEPFILTERS( _ShownDates ),
                    USERELATIONSHIP( {mainDateCol}, {disconnectedDateCol} )
                )
            )
        )
    RETURN
        _Result";
};

string monthsFunc = localPrefix + ".DateControlVisual.PeriodToShow.LastNmonths" + fnSuffix;
bool needMonths = sourceHasMonth && (monthsCheck.Checked || nextMonthsCheck.Checked); // "Next N months" reuses the same function
if (needMonths)
{
    monthsFunc = EnsureLocalFunction(monthsFunc, BuildPeriodFunction("Months", "MONTH"),
        isCalendarSource
            ? $"Shows a measure within a rolling N-month window (blank outside it). Modified from the SQLBI article: added KEEPFILTERS so it also works with Calendar Options, and checks the window with CALCULATETABLE so it also works with monthly/weekly/daily average measures. Uses the '{selectedCalendar.Name}' Calendar Option as a calendar reference (requires the 'Enhanced DAX Time Intelligence' preview feature enabled in Power BI Desktop)."
            : "Shows a measure within a rolling N-month window (blank outside it). Modified from the SQLBI article: added KEEPFILTERS so it also works with Calendar Options, and checks the window with INTERSECT so it also works with monthly/weekly/daily average measures.");
}

// "Previous N weeks" / "Next N weeks" DAX UDF for the selected Date source - only when it has a Week
// time unit. "Next N weeks" reuses the same function (positive numberOfWeeks looks forward).
string weeksFunc = localPrefix + ".DateControlVisual.PeriodToShow.LastNweeks" + fnSuffix;
bool needWeeks = sourceHasWeek && (weeksCheck.Checked || nextWeeksCheck.Checked);
if (needWeeks)
{
    // Uses the disconnected calendar's Calendar Option (a calendar reference, e.g. 'Weekly
    // disconnected') as the DATESINPERIOD source - required for WEEK, which only accepts a calendar
    // reference, not a column reference (only DAY/MONTH/QUARTER/YEAR are valid with a column). Also
    // requires the "Enhanced DAX Time Intelligence" preview feature enabled in Power BI Desktop.
    weeksFunc = EnsureLocalFunction(weeksFunc, BuildPeriodFunction("Weeks", "WEEK"),
        $"Shows a measure within a rolling N-week window (blank outside it). Modified from the SQLBI article: added KEEPFILTERS so it also works with Calendar Options, and checks the window with CALCULATETABLE so it also works with monthly/weekly/daily average measures. Requires the '{selectedCalendar.Name}' Calendar Option with a Week time unit as a calendar reference (requires the 'Enhanced DAX Time Intelligence' preview feature enabled in Power BI Desktop).");
}

// ============================================================================
// STEP 5 - Create the calculation items
// ============================================================================
if (ytdCheck.Checked && firstDayFuncName != null)
{
    AddCalcItem("This year to selected date" + nameSuffix,
        BuildStartEndBody(
            $"{firstDayFuncName}( 0 /* year offset */ )",
            $"{refDateFuncName}()"
        ), "All dates from the start of the year until the last selected date" + nameSuffix);
}

if (entireYearCheck.Checked && firstDayFuncName != null)
{
    AddCalcItem("Show entire selected year" + nameSuffix,
        BuildStartEndBody(
            $"{firstDayFuncName}( 0 /* year offset */ )",
            $"{lastDayFuncName}( 0 /* year offset */ )"
        ), "Shows the entire year for the last selected date" + nameSuffix);
}

if (monthsCheck.Checked && sourceHasMonth)
{
    foreach (var n in ParseNList(monthsBox.Text))
    {
        AddCalcItem($"Previous {n} Months" + nameSuffix,
            $@"{monthsFunc}(
    SELECTEDMEASURE(),
    -{n}
)", $"Shows the last {n} months from the selected date{nameSuffix}");
    }
}

if (weeksCheck.Checked && sourceHasWeek)
{
    foreach (var n in ParseNList(weeksBox.Text))
    {
        AddCalcItem($"Previous {n} Weeks" + nameSuffix,
            $@"{weeksFunc}(
    SELECTEDMEASURE(),
    -{n}
)", $"Shows the last {n} weeks from the selected date{nameSuffix}");
    }
}

if (nextWeeksCheck.Checked && needWeeks)
{
    foreach (var n in ParseNList(nextWeeksBox.Text))
    {
        AddCalcItem($"Next {n} Weeks" + nameSuffix,
            $@"{weeksFunc}(
    SELECTEDMEASURE(),
    {n}
)", $"Shows the next {n} weeks from the selected date{nameSuffix}");
    }
}

if (upToTodayCheck.Checked && lastDayFuncName != null)
{
    AddCalcItem("Show up to today and last year" + nameSuffix,
        BuildStartEndBody(
            $"{lastDayFuncName}( -1 /* year offset */ )",
            $"{refDateFuncName}()"
        ), "Show up to selected date this year and last date last year (such as for incoming balance)" + nameSuffix);
}

if (yearsBackCheck.Checked && firstDayFuncName != null)
{
    foreach (var n in ParseNList(yearsBackBox.Text))
    {
        AddCalcItem($"Year to date and last {n} years" + nameSuffix,
            BuildStartEndBody(
                $"{firstDayFuncName}( -{n} /* year offset */ )",
                $"{refDateFuncName}()"
            ), $"Shows up to the selected date and also the entire previous {n} year(s)" + nameSuffix);
    }
}

if (remainingCheck.Checked && lastDayFuncName != null)
{
    AddCalcItem("Remaining year" + nameSuffix,
        BuildStartEndBody(
            $"{refDateFuncName}()",
            $"{lastDayFuncName}( 0 /* year offset */ )"
        ), "Shows the rest of the year from the last selected date" + nameSuffix);
}

if (nextMonthsCheck.Checked && needMonths)
{
    foreach (var n in ParseNList(nextMonthsBox.Text))
    {
        AddCalcItem($"Next {n} months" + nameSuffix,
            $@"{monthsFunc}(
    SELECTEDMEASURE(),
    {n}
)", $"Shows the next {n} months from the selected date{nameSuffix}");
    }
}

// ============================================================================
// Done
// ============================================================================
string summary = $"Done. Calculation group: '{calcGroupTable.Name}'." + Environment.NewLine +
    $"Date source: {(isCalendarSource ? selectedCalendar.Name : "Classic")}" + Environment.NewLine +
    $"Main date column: {mainDateCol}" + Environment.NewLine +
    $"Disconnected calendar: {disconnectedTable.Name} ({disconnectedDateCol})";
if (warnings.Any())
    summary += Environment.NewLine + Environment.NewLine + "Notes / warnings:" + Environment.NewLine + string.Join(Environment.NewLine, warnings.Select(w => "- " + w));

Info(summary);

// Separate precedence reminder - only shown the first time the calculation group itself is created
// (not on later re-runs/updates), since Precedence is only ever set here, once.
if (isNewCalcGroupTable)
{
    int precedence = calcGroupTable.CalculationGroup.Precedence;
    var conflicting = Model.CalculationGroups
        .Where(cg => cg != calcGroupTable && cg.CalculationGroup.Precedence == precedence)
        .Select(cg => cg.Name)
        .ToList();

    string precedenceMsg = $"Calculation group '{calcGroupTable.Name}' was created with Precedence = {precedence}.";
    if (conflicting.Any())
        precedenceMsg += Environment.NewLine + Environment.NewLine +
            $"WARNING: Precedence {precedence} is already used by: {string.Join(", ", conflicting)}. Two calculation groups sharing the same Precedence give an undefined/unreliable evaluation order - change one of them.";
    precedenceMsg += Environment.NewLine + Environment.NewLine +
        "Please review and adjust Precedence so it is set HIGHER than any Time Intelligence calculation group (applied after it), and typically LOWER than e.g. a Unit Conversion or Currency Conversion calculation group (applied last).";

    Info(precedenceMsg);
}

