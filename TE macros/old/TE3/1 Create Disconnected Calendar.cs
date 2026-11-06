
/*
 * Title: Create Disconnected Calendar Table
 * 
 * Author: Eivind Haugen
 * 
 * This script, when executed, will let you select a calendar table to copy.
 * It will write a new table (currently for XLMA-endpoint (Premium) or PBIP. If not, it will create a calculated table
 * The purpose of this table is to use it as part of a calculation group months to show, to dynamically select axis in visuals
 * There is a separate macro to format it and connect to the main calendar-table. And it can use standard format if a standard calendar table is used a source
 * Using write-operation, it will require an update if using PBIP, before applying remaining macros
 */





#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System;
using System.Drawing;
using System.Windows.Forms;

// Hide script execution dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Get all tables in the model
var _dateTables = Model.Tables.ToList();

// Filter only tables with the data category "Time"
var filteredDateTables = _dateTables.Where(t => t.DataCategory == "Time").ToList();

// If no tables with "Time" data category are found, use the original list
if (filteredDateTables.Count > 0)
{
    _dateTables = filteredDateTables;
}

// Select the table to reference
var _calendarTable = SelectTable(_dateTables, null, "Select table to be referenced in the disconnected calendar");

if (_calendarTable == null)
{
    Info("No table selected. Operation aborted.");
    return;
}

// Remove single quotes from the table name
string _cleanedCalendarTable = _calendarTable.DaxObjectFullName.Trim('\'');

// Prompt user for the name of the new disconnected calendar table
string _disconnectedCalendarTableName = Interaction.InputBox("Provide the name of the new calendar table", "Table Name", "Disconnected_Calendar");

// Check if the table already exists
if (Model.Tables.Contains(_disconnectedCalendarTableName))
{
    Info($"Table '{_disconnectedCalendarTableName}' already exists. Skipping creation.");
}
else
{
    try
    {
        // Attempt to create a new table
        var _Table = Model.AddTable(_disconnectedCalendarTableName, false);

        // Define the M expression dynamically using the selected source table
        string _MPartition = $@"
let
    Source = #""{_cleanedCalendarTable}""
in
    Source";

        // Add partition with dynamic source table
        var partition = _Table.AddMPartition(_disconnectedCalendarTableName, _MPartition);
        partition.Mode = ModeType.Import; // Set refresh policy to Import



        Info($"Created '{_disconnectedCalendarTableName}' table successfully, using '{_calendarTable.Name}' as the source.");
    }
    catch (Exception ex)
    {
        Info($"Write-operation not supported. Creating calculated table instead.\n\nReason: {ex.Message}");

        // Create a calculated table as a fallback
        string calculatedTableExpression = _calendarTable.DaxObjectFullName;
        var calcTable = Model.AddCalculatedTable(_disconnectedCalendarTableName, calculatedTableExpression);

        Info($"Calculated table '{_disconnectedCalendarTableName}' created successfully.");
    }
}
