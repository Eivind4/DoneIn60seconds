// Credit to Fernan for original DAX idea
// Converted to C# Script from DAX script created by David Kofod Hanna

// Check if the calculated table already exists, if so delete it
var existingTable = Model.Tables.FirstOrDefault(t => t.Name == "Model Documentation");
if (existingTable != null)
{
    existingTable.Delete();
}

// DAX to create the calculated table
var docTableDax = @"
VAR _columns =
    SELECTCOLUMNS(
        FILTER(
            INFO.VIEW.COLUMNS(),
            [Table] <> ""Model Documentation"" && NOT([IsHidden])
        ),
        ""Type"", ""Column"",
        ""Name"", [Name],
        ""Description"", [Description],
        ""Location"", [Table],
        ""Expression"", [Expression]
    )
VAR _measures =
    SELECTCOLUMNS(
        FILTER(
            INFO.VIEW.MEASURES(),
            [Table] <> ""Model Documentation"" && NOT([IsHidden])
        ),
        ""Type"", ""Measure"",
        ""Name"", [Name],
        ""Description"", [Description],
        ""Location"", [Table],
        ""Expression"", [Expression]
    )
VAR _tables =
    SELECTCOLUMNS(
        FILTER(
            INFO.VIEW.TABLES(),
            [Name] <> ""Model Documentation"" && [Name] <> ""Calculations"" && NOT([IsHidden])
        ),
        ""Type"", ""Table"",
        ""Name"", [Name],
        ""Description"", [Description],
        ""Location"", BLANK(),
        ""Expression"", [Expression]
    )
VAR _relationships =
    SELECTCOLUMNS(
        INFO.VIEW.RELATIONSHIPS(),
        ""Type"", ""Relationship"",
        ""Name"", [Relationship],
        ""Description"", BLANK(),
        ""Location"", BLANK(),
        ""Expression"", [Relationship]
    )
RETURN
    UNION(_columns, _measures, _tables, _relationships)
";

// Create the new calculated table
var documentationTable = Model.AddCalculatedTable("Model Documentation", docTableDax);

// Create the measures
var measure1 = documentationTable.AddMeasure(
    "# of Measures",
    "CALCULATE(COUNTROWS('Model Documentation'), 'Model Documentation'[Type] = \"Measure\")"
);
measure1.FormatString = "#,0";
measure1.Description = "COUNTROWS('Model Documentation') where Type = 'Measure'";

var measure2 = documentationTable.AddMeasure(
    "# of Columns",
    "CALCULATE(COUNTROWS('Model Documentation'), 'Model Documentation'[Type] = \"Column\")"
);
measure2.FormatString = "#,0";
measure2.Description = "COUNTROWS('Model Documentation') where Type = 'Column'";

var measure3 = documentationTable.AddMeasure(
    "# of Tables",
    "CALCULATE(COUNTROWS('Model Documentation'), 'Model Documentation'[Type] = \"Table\")"
);
measure3.FormatString = "#,0";
measure3.Description = "COUNTROWS('Model Documentation') where Type = 'Table'";

var measure4 = documentationTable.AddMeasure(
    "# of Relationship",
    "CALCULATE(COUNTROWS('Model Documentation'), 'Model Documentation'[Type] = \"Relationship\")"
);
measure4.FormatString = "#,0";
measure4.Description = "COUNTROWS('Model Documentation') where Type = 'Relationship'";

// Optional: Refresh schema (Tabular Editor 3 only)
// Model.Sync();
