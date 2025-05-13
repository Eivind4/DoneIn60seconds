// This is a modification of the auto-generate script provided by Tabular Editor.
// It is modified such that the user can select the table where the new measure is to be added
// It adds logic setting AvailableInMDX = False for the base column

//
 // Title: Auto-generate SUM measures from columns
 // 
 // Author: Daniel Otykier, twitter.com/DOtykier
 // 
 // This script, when executed, will loop through the currently selected columns,
 // creating one SUM measure for each column and also hiding the column itself.
 //


using System.Windows.Forms;
using System.Linq;

// Prompt: Add to custom table?
var result = MessageBox.Show(
    "Do you want to add the new measure(s) to a custom table?\n\nClick 'Yes' to choose a table.\nClick 'No' to add to each column's original table.",
    "Choose Target Table",
    MessageBoxButtons.YesNoCancel
);

if (result == DialogResult.Cancel)
{
    return; // User cancelled
}

Table selectedTable = null;

if (result == DialogResult.Yes)
{
    // Build a form with a ListBox for table selection
    Form form = new Form();
    ListBox listBox = new ListBox();
    Button okButton = new Button();
    Button cancelButton = new Button();

    form.Text = "Select a Table";
    form.Width = 400;
    form.Height = 300;
    form.FormBorderStyle = FormBorderStyle.FixedDialog;
    form.StartPosition = FormStartPosition.CenterScreen;
    form.MinimizeBox = false;
    form.MaximizeBox = false;

    listBox.Width = 360;
    listBox.Height = 200;
    listBox.Left = 10;
    listBox.Top = 10;
    listBox.SelectionMode = SelectionMode.One;
    listBox.DataSource = Model.Tables.Select(t => t.Name).ToList();

    okButton.Text = "OK";
    okButton.Left = 220;
    okButton.Top = 220;
    okButton.DialogResult = DialogResult.OK;

    cancelButton.Text = "Cancel";
    cancelButton.Left = 300;
    cancelButton.Top = 220;
    cancelButton.DialogResult = DialogResult.Cancel;

    form.Controls.Add(listBox);
    form.Controls.Add(okButton);
    form.Controls.Add(cancelButton);

    form.AcceptButton = okButton;
    form.CancelButton = cancelButton;

    var dialogResult = form.ShowDialog();

    if (dialogResult == DialogResult.OK && listBox.SelectedItem != null)
    {
        string tableName = listBox.SelectedItem.ToString();
        if (!string.IsNullOrEmpty(tableName) && Model.Tables.Contains(tableName))
        {
            selectedTable = Model.Tables[tableName];
        }
    }
    else
    {
        Info("No table selected. Measures will be added to the column's original table.");
    }
}

// Loop through selected columns and add SUM measures
foreach (var c in Selected.Columns)
{
    var finalTable = selectedTable ?? c.Table;
    var daxExpression = "SUM(" + c.DaxObjectFullName + ")";

    bool measureExists = Model.AllMeasures.Any(m =>
        m.Expression.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") == daxExpression.Replace(" ", "")
    );

    if (!measureExists)
    {
        var measure = finalTable.AddMeasure(
            "Total " + c.Name,
            daxExpression,
            c.Table.Name
        );

        measure.FormatString = "#,0";
        measure.Description = "This measure is the sum of column " + c.DaxObjectFullName;

        c.IsHidden = true;
        c.IsAvailableInMDX = false;

        c.SummarizeBy = AggregateFunction.None;
    }
    else
    {
        Output("Skipped: Measure already exists for column '" + c.Name + "' with expression '" + daxExpression + "'");
    }
}

