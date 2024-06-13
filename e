let
    // Load the source data from SharePoint
    Source = SharePoint.Files("https://nabrookfieldgis.sharepoint.com/sites/rcRp1WsPilot", [ApiVersion = 15]),
    
    // Filter rows based on folder path and file extension
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Folder Path], "Jun 22 onwards")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each ([Extension] = ".xlsx")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),
    
    // Invoke custom function to transform file
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (2)", each #"Transform File (2)"([Content])),
    
    // Rename and expand columns
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (2)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (2)", Table.ColumnNames(#"Transform File (2)"(#"Sample File (2)"))),

    // Filter out non-month values
    #"Filtered Rows3" = Table.SelectRows(#"Expanded Table Column1", each Text.Contains([Month], "January") or Text.Contains([Month], "February") or Text.Contains([Month], "March") or Text.Contains([Month], "April") or Text.Contains([Month], "May") or Text.Contains([Month], "June") or Text.Contains([Month], "July") or Text.Contains([Month], "August") or Text.Contains([Month], "September") or Text.Contains([Month], "October") or Text.Contains([Month], "November") or Text.Contains([Month], "December")),

    // Ensure the "Month" column is consistent and convert to text
    #"Transformed Month Column" = Table.TransformColumns(#"Filtered Rows3", {{"Month", each Text.From(_), type text}}),
    
    #"Added Index" = Table.AddIndexColumn(#"Transformed Month Column", "Index", 0, 1),
    #"Inserted Addition" = Table.AddColumn(#"Added Index", "Addition", each [Index] + 1, type number),

    // Ensure the "Month" column exists before merging
    #"Check Inserted Addition" = Table.SelectColumns(#"Inserted Addition", {"Month", "Index", "Addition"}),

    // Perform the nested join
    #"Merged Queries" = Table.NestedJoin(#"Inserted Addition", {"Index"}, #"Inserted Addition", {"Addition"}, "Inserted Addition", JoinKind.LeftOuter),

    // Inspect the structure of the merged queries
    #"Check Merged Queries Structure" = Table.SelectColumns(#"Merged Queries", {"Inserted Addition"}),

    // Ensure the "Month" column exists in the nested table before expanding
    #"Check Merged Queries" = Table.SelectColumns(#"Merged Queries", {"Inserted Addition"}),

    // Expand the nested table and ensure the "Month" column is included
    #"Expanded Insert Addition" = Table.ExpandTableColumn(#"Merged Queries", "Inserted Addition", {"Month"}, {"Inserted Addition.Month"}),

    // Further transformations
    #"Filtered Rows4" = Table.SelectRows(#"Expanded Insert Addition", each [Inserted Addition.Month] <> null and not Text.Contains([Inserted Addition.Month], ",")),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows4", each ([Inserted Addition.Month] <> "(Rel: 10.0.0.0)")),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows2",{{"Inserted Addition.Month", "Queue"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns",{{"Month", type date}}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Changed Type2",{{"Avg", type text}, {"GOS1", Percentage.Type}, {"GOS2", Percentage.Type}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Changed Type3",{{"Min", "Logged on Users Min"}, {"Addition", "Addition Index"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns4", "Custom", each if Text.StartsWith([Queue], "Queue 5") then "Client E-mail" else if Text.Contains([Queue], "Vendor") then "Vendor" else "Client Phone"),
    #"Renamed Columns2" = Table.RenameColumns(#"Added Conditional Column",{{"Offered", "Total Offered"}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Renamed Columns2",{{"Amt", "Amt Handled this Line"}, {"Avg", "Avg Handled this Line"}, {"Lngst", "Lngst Handled this Line"}, {"Amt_1", "Amt Handled other Line"}, {"Lngst_3", "Lngst Handled other Line"}, {"Amt_4", "Amt Ended"}, {"Abdnd", "Amt Abdnd"}, {"Custom", "Line"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns3",{{"Total Offered", Int64.Type}, {"Amt Handled this Line", Int64.Type}, {"Amt Handled other Line", Int64.Type}, {"Amt Ended", Int64.Type}, {"Amt Abdnd", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Amount Abnd %", each [Amt Abdnd]/[Total Offered]),
    #"Changed Type4" = Table.TransformColumnTypes(#"Added Custom",{{"Amount Abnd %", Percentage.Type}})
in
    #"Changed Type4"
