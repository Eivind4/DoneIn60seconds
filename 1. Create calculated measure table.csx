using System;
using System.Windows.Forms;
using System.Linq;

public class InputDialog
{
    public static string ShowInputDialog(string text, string caption, string defaultValue)
    {
        Form prompt = new Form()
        {
            Width = 600,
            Height = 150,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = caption,
            StartPosition = FormStartPosition.CenterScreen
        };
        Label textLabel = new Label() { Left = 50, Top = 20, Text = text, AutoSize = true };
        TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 500, Text = defaultValue };
        Button confirmation = new Button() { Text = "OK", Left = 450, Width = 100, Top = 70, DialogResult = DialogResult.OK };
        confirmation.Click += (sender, e) => { prompt.Close(); };
        prompt.Controls.Add(textLabel);
        prompt.Controls.Add(textBox);
        prompt.Controls.Add(confirmation);
        prompt.AcceptButton = confirmation;

        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
    }
}

Application.UseWaitCursor = false;

string defaultTableName = "Model Measures";
string tableName = InputDialog.ShowInputDialog("Enter name for the measure table:", "Create Measure Table", defaultTableName);

if (string.IsNullOrWhiteSpace(tableName))
{
    Info("Table name was not provided. Operation cancelled.");
    return;
}

// Check if table already exists
if (Model.Tables.Any(t => t.Name == tableName))
{
    Info("A table named '" + tableName + "' already exists. No action taken.");
    return;
}

try
{
    string daxTable = "DATATABLE(\"MeasureName\", STRING, {{\"Default Measure\"}})";
    var newTable = Model.AddCalculatedTable(tableName, daxTable);

    if (newTable != null)
    {
        newTable.Description = "Table to collect measures";
        foreach (var col in newTable.Columns)
        {
            col.IsHidden = true;
        }

        //  Info("Successfully created calculated table '" + tableName + "' for storing measures.");
    }
    else
    {
        Info("Failed to create the calculated table '" + tableName + "'.");
    }
}
catch (Exception ex)
{
    Info("Error during calculated table creation:\n\n" + ex.Message);
}
