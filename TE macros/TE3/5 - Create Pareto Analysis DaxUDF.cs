// ADD PARETO ANALYSIS DAX UDFs (MODEL-INDEPENDENT PART)
// =========================================================
//
// Purpose: Generate the reusable, model-INDEPENDENT DAX User Defined Functions (UDFs) behind the
// "Dynamic Pareto Analysis" pattern described by SQLBI:
//   https://www.sqlbi.com/articles/dynamic-pareto-analysis-in-power-bi/
//
// This macro deliberately covers only the generic/reusable building blocks - not a complete,
// ready-to-use Pareto report. Per the requester's scope:
//   - The "Metric" table + "Metric Value" dynamic-metric-selection pattern from the article is
//     SKIPPED here (add it separately if/when you want dynamic metric switching).
//   - The tooltip-oriented "Pareto Amount" / "Customer Name" measures from the article (the "last
//     measure in the Sales table" in the article) are SKIPPED - these are highly model-specific
//     (they hard-code the key/name columns) and should be built later, tailored to the model.
//   - No "Local." model-dependent wrapper functions are created by this macro (unlike the other
//     DaxUDF generator macros) - only the generic functions below. Wrap these with a follow-up
//     macro/manual step once you know the actual dimension table, key column, and metric(s).
//
// Functions created (namespace: ParetoAnalysis, with ParetoAnalysis.VisualCalc for the two
// functions intended to be used from inside a Visual calculation, since they rely on the ROWS
// keyword which is only meaningful in that context):
//
//   - ParetoAnalysis.XAxis(baseTable)
//       Generic version of the article's "X_axis" calculated table, parameterized so any
//       dimension table can be used instead of the article's hard-coded "Customer" table:
//         X_axis = ParetoAnalysis.XAxis(Customer)
//
//   - ParetoAnalysis.Cumulative(keyColumn, metricValue, xAxisColumn)
//       Generic version of the article's "Pareto Cumulative" measure, parameterized so the key
//       column (article: Customer[CustomerKey]), the metric (article: [Metric Value]), and the
//       X-axis column (article: X_axis[X]) are all supplied by the caller instead of hard-coded.
//
//   - ParetoAnalysis.Pct(paretoCumulative, metricValue)
//       Generic version of the article's "Pareto %" measure, parameterized to take the two
//       measures being divided (article: [Pareto Cumulative] and [Metric Value]) as inputs.
//
//   - ParetoAnalysis.VisualCalc.Pct(runningSum)
//       Generic version of the article's "%" Visual calculation, which divides the running sum by
//       its COLLAPSE(..., ROWS) total. Use with a manually-created "Running sum" Visual calculation
//       (see note below on why RunningSum/RunningSumPerCategory could NOT be turned into UDFs).
//
// NOTE - "Running sum" / "Running sum per category" Visual calculations are NOT included as UDFs:
//   WINDOW's ORDERBY/PARTITIONBY require a direct column reference to resolve lineage within the
//   Visual calculation's ROWS relation. When metricValue/partitionColumn are passed in as function
//   parameters, that lineage is lost and DAX raises "Column [...] not found in the Relation
//   argument of the WINDOW function." There is no reliable workaround for this - so these two
//   Visual calculations must still be authored directly (not via a UDF) using the article's
//   pattern, substituting your own metric/partition column:
//     Running sum =
//         SUMX(
//             WINDOW(1, ABS, 0, REL, ROWS, ORDERBY([<your metric>], DESC)),
//             [<your metric>]
//         )
//   ParetoAnalysis.VisualCalc.Pct(runningSum) can still be used on top of that Visual calculation.
//
// Safety checks performed before creating each function:
// - If a function with an IDENTICAL DAX expression already exists under a different name, the new
//   UDF is NOT created - the existing function is reused/referenced instead.
// - If a function with the SAME NAME already exists:
//     - and its expression is identical, nothing happens (already up to date, left unchanged).
//     - and its expression differs, the user is asked to Rename the new one, Overwrite the
//       existing (outdated) definition, or Skip creating it (never silently overwritten).
//
#r "System.Drawing"
using System.Windows.Forms;
using System.Drawing;

Application.UseWaitCursor = false;
WaitFormVisible = false;

const string articleUrl = "https://www.sqlbi.com/articles/dynamic-pareto-analysis-in-power-bi/";

Func<string, string> NormalizeExpr = s => string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));

// ================================================================================================
// DIALOG: Confirm before creating - lists exactly what will be (re)created, with a link to the
// source article for reference.
// ================================================================================================
var introForm = new Form
{
    Text = "Pareto Analysis DaxUDF - model-independent functions",
    MaximizeBox = false,
    MinimizeBox = false,
    ShowIcon = false,
    AutoSize = true,
    AutoSizeMode = AutoSizeMode.GrowAndShrink,
    Width = 560,
    FormBorderStyle = FormBorderStyle.FixedDialog,
    StartPosition = FormStartPosition.CenterScreen
};
var introFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
introForm.Controls.Add(introFlow);

introFlow.Controls.Add(new Label
{
    Text = "This will create the following model-independent DAX UDFs, based on:\r\n" + articleUrl,
    AutoSize = true,
    MaximumSize = new Size(520, 0),
    Margin = new Padding(0, 0, 0, 10)
});

Action<string> bullet = s => introFlow.Controls.Add(new Label { Text = "- " + s, AutoSize = true, MaximumSize = new Size(520, 0), Margin = new Padding(0, 0, 0, 6) });
bullet("ParetoAnalysis.XAxis(baseTable) - numeric axis table (replaces the article's Customer table)");
bullet("ParetoAnalysis.Cumulative(keyColumn, metricValue, xAxisColumn) - Pareto Cumulative");
bullet("ParetoAnalysis.Pct(paretoCumulative, metricValue) - Pareto %");
bullet("ParetoAnalysis.VisualCalc.Pct(runningSum) - Visual calculation Running sum % (use on a manually-authored Running sum Visual calc - see header note)");

introFlow.Controls.Add(new Label
{
    Text = "Not included (per scope): the Metric/Metric Value dynamic-metric pattern, and the tooltip-only Pareto Amount / Customer Name measures - add these later, tailored to your model.",
    AutoSize = true,
    MaximumSize = new Size(520, 0),
    Margin = new Padding(0, 10, 0, 10)
});

var introBtnPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 5, 0, 0) };
var introOk     = new Button { Text = "Create",  AutoSize = true, DialogResult = DialogResult.OK,     Margin = new Padding(0, 0, 8, 0) };
var introCancel = new Button { Text = "Cancel",  AutoSize = true, DialogResult = DialogResult.Cancel };
introBtnPanel.Controls.Add(introOk);
introBtnPanel.Controls.Add(introCancel);
introFlow.Controls.Add(introBtnPanel);
introForm.AcceptButton = introOk;
introForm.CancelButton = introCancel;

if (introForm.ShowDialog() == DialogResult.Cancel) return;

// ================================================================================================
// Create the functions - same skip/rename/overwrite conflict handling used by the other DaxUDF
// generator macros.
// ================================================================================================
var notes = new List<string>();

Action<string, string, string> CreateUDF = (funcName, daxBody, description) =>
{
    var normalizedNew = NormalizeExpr(daxBody);

    var dupExprFunc = Model.Functions.FirstOrDefault(f => f.Name != funcName && NormalizeExpr(f.Expression) == normalizedNew);
    if (dupExprFunc != null)
    {
        notes.Add($"'{funcName}' was NOT created - using the existing function '{dupExprFunc.Name}' instead, since it has an identical expression.");
        return;
    }

    var existing = Model.Functions.FirstOrDefault(f => f.Name == funcName);
    if (existing != null)
    {
        if (NormalizeExpr(existing.Expression) == normalizedNew)
        {
            notes.Add($"'{funcName}' already exists with the same expression - left unchanged.");
            return;
        }

        var conflictDlg = new Form { Text = "Function name already exists", Width = 480, Height = 250, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false };
        var conflictFlow = new FlowLayoutPanel { Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        conflictDlg.Controls.Add(conflictFlow);
        conflictFlow.Controls.Add(new Label { Text = $"A function named '{funcName}' already exists with a different expression.\r\nChoose how to proceed:", AutoSize = true, MaximumSize = new Size(440, 0), Margin = new Padding(0, 0, 0, 10) });

        var renameRow   = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
        var renameRadio = new RadioButton { Text = "Rename to:", AutoSize = true, Checked = true, Margin = new Padding(0, 3, 8, 0) };
        var renameBox   = new TextBox { Width = 250, Text = funcName };
        renameRow.Controls.Add(renameRadio);
        renameRow.Controls.Add(renameBox);
        conflictFlow.Controls.Add(renameRow);

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

        if (overwriteRadio.Checked)
        {
            existing.Expression  = daxBody;
            existing.Description = description;
            notes.Add($"'{funcName}' already existed with a different expression - its definition was OVERWRITTEN with the new one.");
            return;
        }
        else if (skipRadio.Checked)
        {
            notes.Add($"'{funcName}' was NOT created - a function with that name already exists (existing definition kept).");
            return;
        }
        else // renameRadio
        {
            string newName = renameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName) || newName == funcName)
            {
                notes.Add($"'{funcName}' was NOT created - no valid new name was provided (existing definition kept).");
                return;
            }
            if (Model.Functions.Any(f => f.Name == newName))
            {
                notes.Add($"'{funcName}' was NOT created - the alternate name '{newName}' also already exists.");
                return;
            }
            funcName = newName;
        }
    }

    var fn = Model.AddFunction(funcName);
    fn.Expression  = daxBody;
    fn.Description = description;
};

// ------------------------------------------------------------------------------------------------
// ParetoAnalysis.XAxis - replaces the article's hard-coded "Customer" table with a parameter.
// Usage (model-dependent step, not part of this macro): X_axis = ParetoAnalysis.XAxis(Customer)
// ------------------------------------------------------------------------------------------------
string xAxisBody = $@"(
    // {articleUrl}
    // The table whose row count defines the numeric axis range (article uses Customer)
    baseTable: ANYREF
) =>
    SELECTCOLUMNS(
        GENERATESERIES(1, COUNTROWS(baseTable)),
        ""X"", [Value]
    )";
string xAxisDescription = $"Dynamic Pareto analysis - numeric continuous axis table (1..row count of baseTable), replacing the article's hard-coded Customer table. Use as a calculated table, e.g. X_axis = ParetoAnalysis.XAxis(Customer). Reference: {articleUrl}";

// ------------------------------------------------------------------------------------------------
// ParetoAnalysis.Cumulative - replaces Customer[CustomerKey], [Metric Value], and X_axis[X] with
// parameters (keyColumn, metricValue, xAxisColumn).
// ------------------------------------------------------------------------------------------------
string cumulativeBody = $@"(
    // {articleUrl}
    // The key column uniquely identifying each item to rank (article: Customer[CustomerKey])
    keyColumn: ANYREF,
    // The measure used to rank and sum items (article: [Metric Value])
    metricValue: ANYREF,
    // The numeric axis column from the disconnected X-axis table (article: X_axis[X])
    xAxisColumn: ANYREF
) =>
    VAR _XPos = SELECTEDVALUE(xAxisColumn)
    VAR _Points =
        ADDCOLUMNS(
            VALUES(keyColumn),
            ""@Value"", metricValue
        )
    VAR _CumulatedPoints =
        WINDOW(
            1, ABS,
            _XPos, ABS,
            _Points,
            ORDERBY([@Value], DESC, keyColumn, ASC)
        )
    VAR _Result =
        SUMX(_CumulatedPoints, [@Value])
    RETURN
        _Result";
string cumulativeDescription = $"Dynamic Pareto analysis - cumulative value up to the position selected on the numeric X-axis. Parameterized version of the article's 'Pareto Cumulative' measure (keyColumn = Customer[CustomerKey], metricValue = [Metric Value], xAxisColumn = X_axis[X] in the article). Reference: {articleUrl}";

// ------------------------------------------------------------------------------------------------
// ParetoAnalysis.Pct - replaces [Pareto Cumulative] and [Metric Value] with measure inputs.
// ------------------------------------------------------------------------------------------------
string pctBody = $@"(
    // {articleUrl}
    // The cumulative Pareto value (typically built with ParetoAnalysis.Cumulative)
    paretoCumulative: ANYREF,
    // The total metric value being distributed (article: [Metric Value])
    metricValue: ANYREF
) =>
    DIVIDE(paretoCumulative, metricValue)";
string pctDescription = $"Dynamic Pareto analysis - cumulative value expressed as a percentage of the total, for plotting against the numeric X-axis. Parameterized version of the article's 'Pareto %' measure. Reference: {articleUrl}";

// ------------------------------------------------------------------------------------------------
// ParetoAnalysis.VisualCalc.Pct - replaces [Running sum] with a parameter. Applies to a manually-
// authored "Running sum" Visual calculation (see header note - WINDOW/ORDERBY lineage prevents
// the running sum itself from being a UDF). COLLAPSE(runningSum, ROWS) does not require column
// lineage the way WINDOW's ORDERBY/PARTITIONBY do, so this one works fine as a UDF.
// ------------------------------------------------------------------------------------------------
string runningSumPctBody = $@"(
    // {articleUrl}
    // The running sum Visual calculation result (a manually-authored 'Running sum' measure)
    runningSum: ANYREF
) =>
    DIVIDE(
        runningSum,
        COLLAPSE(runningSum, ROWS)
    )";
string runningSumPctDescription = $"Dynamic Pareto analysis - Visual calculation percentage of total, dividing runningSum by its COLLAPSE(..., ROWS) total. Parameterized version of the article's '%' Visual calculation. Use directly as a Visual calculation formula on top of a manually-authored 'Running sum' Visual calculation (running sum itself cannot be a UDF - see header note). Reference: {articleUrl}";

CreateUDF("ParetoAnalysis.XAxis", xAxisBody, xAxisDescription);
CreateUDF("ParetoAnalysis.Cumulative", cumulativeBody, cumulativeDescription);
CreateUDF("ParetoAnalysis.Pct", pctBody, pctDescription);
CreateUDF("ParetoAnalysis.VisualCalc.Pct", runningSumPctBody, runningSumPctDescription);

// ================================================================================================
// Done
// ================================================================================================
string doneSummary = "Done! Created/updated the model-independent Pareto Analysis functions.\n\n" +
    "Next steps (model-dependent, not part of this macro):\n" +
    "- Create the X_axis calculated table, e.g.: X_axis = ParetoAnalysis.XAxis(<your dimension table>)\n" +
    "- Build a 'Pareto Cumulative' measure calling ParetoAnalysis.Cumulative(<keyColumn>, <metric measure>, X_axis[X])\n" +
    "- Build a 'Pareto %' measure calling ParetoAnalysis.Pct([Pareto Cumulative], <metric measure>)\n" +
    "- Manually author a 'Running sum' Visual calculation (WINDOW/ORDERBY on your metric - see header note for why this can't be a UDF), then apply ParetoAnalysis.VisualCalc.Pct(runningSum) on top of it\n" +
    "- Optionally add the Metric/Metric Value dynamic-metric pattern and the tooltip-only Pareto Amount / Customer Name measures later, tailored to your model.";
if (notes.Any())
    doneSummary += Environment.NewLine + Environment.NewLine + "Notes:" + Environment.NewLine + string.Join(Environment.NewLine, notes.Select(n => "- " + n));

Info(doneSummary);
