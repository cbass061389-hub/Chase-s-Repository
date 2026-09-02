shared #"Forecast Qty" = let
    Source = Csv.Document(
        Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2606291&c=3492685&h=<REDACTED>&_xt=.csv"),
        [Delimiter=",", Columns=42, Encoding=65001, QuoteStyle=QuoteStyle.None]
    ),

    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    #"Changed Type" = Table.TransformColumnTypes(
        #"Promoted Headers",
        {
            {"Item", type text},
            {"Internal ID", Int64.Type},
            {"Product Category", type text},
            {"Territory", type text},
            {"Sub Location", type text},
            {"Lead Group", type text},
            {"Forecast QTY 01-2025", type text},
            {"Forecast QTY 02-2025", type text},
            {"Forecast QTY 03-2025", type text},
            {"Forecast QTY 04-2025", type text},
            {"Forecast QTY 05-2025", type text},
            {"Forecast QTY 06-2025", Int64.Type},
            {"Forecast QTY 07-2025", Int64.Type},
            {"Forecast QTY 08-2025", Int64.Type},
            {"Forecast QTY 09-2025", Int64.Type},
            {"Forecast QTY 10-2025", Int64.Type},
            {"Forecast QTY 11-2025", Int64.Type},
            {"Forecast QTY 12-2025", Int64.Type},
            {"Forecast QTY 01-2026", Int64.Type},
            {"Forecast QTY 02-2026", Int64.Type},
            {"Forecast QTY 03-2026", Int64.Type},
            {"Forecast QTY 04-2026", Int64.Type},
            {"Forecast QTY 05-2026", Int64.Type},
            {"Forecast QTY 06-2026", Int64.Type},
            {"Forecast QTY 07-2026", Int64.Type},
            {"Forecast QTY 08-2026", Int64.Type},
            {"Forecast QTY 09-2026", Int64.Type},
            {"Forecast QTY 10-2026", Int64.Type},
            {"Forecast QTY 11-2026", Int64.Type},
            {"Forecast QTY 12-2026", Int64.Type},
            {"Forecast QTY 01-2027", Int64.Type},
            {"Forecast QTY 02-2027", Int64.Type},
            {"Forecast QTY 03-2027", Int64.Type},
            {"Forecast QTY 04-2027", Int64.Type},
            {"Forecast QTY 05-2027", Int64.Type},
            {"Forecast QTY 06-2027", Int64.Type},
            {"Forecast QTY 07-2027", Int64.Type},
            {"Forecast QTY 08-2027", Int64.Type},
            {"Forecast QTY 09-2027", Int64.Type},
            {"Forecast QTY 10-2027", Int64.Type},
            {"Forecast QTY 11-2027", Int64.Type},
            {"Forecast QTY 12-2027", Int64.Type}
        }
    ),

    // Turn monthly columns into rows
    #"Unpivoted Other Columns" =
        Table.UnpivotOtherColumns(
            #"Changed Type",
            {"Lead Group", "Sub Location", "Territory", "Product Category", "Internal ID", "Item"},
            "Attribute",
            "Value"
        ),
    // Keep only the "MM-YYYY" piece and convert to date
    #"Extracted Last Characters" =
        Table.TransformColumns(
            #"Unpivoted Other Columns",
            {{"Attribute", each Text.End(_, 7), type text}}
        ),

    #"Changed Type1" =
        Table.TransformColumnTypes(#"Extracted Last Characters", {{"Attribute", type date}}),

    #"Renamed Columns" =
        Table.RenameColumns(#"Changed Type1", {{"Value", "Forecast Qty"}}),

    #"Changed Type2" =
        Table.TransformColumnTypes(#"Renamed Columns", {{"Forecast Qty", Int64.Type}}),

    // Build SKU from Item (same as you had)
    #"Added Custom" =
        Table.AddColumn(
            #"Changed Type2",
            "SKU",
            each if Text.Contains([Item], ":") then Text.AfterDelimiter([Item], ":") else [Item]
        ),

    #"Reordered Columns" =
        Table.ReorderColumns(
            #"Added Custom",
            {"SKU", "Item", "Internal ID", "Product Category", "Territory", "Sub Location", "Lead Group", "Attribute", "Forecast Qty"}
        ),

    #"Removed Columns" =
        Table.RemoveColumns(#"Reordered Columns", {"Item"}),

    // Extract Year/Month from the Attribute date
    #"Added Year" =
        Table.AddColumn(#"Removed Columns", "Year", each Date.Year([Attribute]), Int64.Type),

    #"Added Month" =
        Table.AddColumn(#"Added Year", "Month", each Date.Month([Attribute]), Int64.Type),

    // Aggregate to monthly SKU × Sub Location
    #"Grouped Rows1" =
        Table.Group(#"Added Month", {"SKU", "Sub Location", "Year", "Month", "Product Category", "Lead Group", "Territory"}, {{"Forecast Qty", each List.Sum([Forecast Qty]), type nullable number}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Grouped Rows1",{"SKU", "Sub Location", "Year", "Month", "Forecast Qty", "Product Category"}),

    // Map Sub Location → Regional ID (mirroring Forecast query logic)
    #"Added Regional ID" =
        Table.AddColumn(#"Reordered Columns2", "Regional ID", each if [Lead Group] = "Amazon" then "Amazon" else if [Lead Group] = "PBS Events" then "PBS Events" else if [Lead Group] = "eComm" then "eComm"
                else if [Lead Group] = "Sponsorships Representing" or [Lead Group] = "Sponsorships Product Compensation" then "Sponsorship"
                else if [Lead Group] = "Tradeshows" then "Tradeshows"
                else if [Lead Group] = "Beckmann" then "EMEA"
                else if [Sub Location] = "Americas" then "Americas"
                else if [Sub Location] = "Asia" then "Asia"
                else if [Sub Location] = "EMEA" then "EMEA"
                else null),

    #"Changed Type3" =
        Table.TransformColumnTypes(
            #"Added Regional ID",
            {{"Regional ID", type text}, {"Forecast Qty", Int64.Type}}
        ),

    // Build YearMonth like in the Forecast query
    #"Added YearMonth" =
        Table.AddColumn(#"Changed Type3", "YearMonth", each Text.From([Year])& Text.PadStart(Text.From([Month]), 2, "0")),

    // REVO Region mirroring the template
    #"Added REVO Region" =
        Table.AddColumn(
            #"Added YearMonth",
            "REVO Region",
            each
                if [Regional ID] = "Americas" then "Americas"
                else if [Regional ID] = "EMEA" then "EMEA"
                else if [Regional ID] = "AP" then "AP"
                else if [Regional ID] = "B2C" then "B2C"
                else null
        ),

    #"Added Custom1" = Table.AddColumn(#"Added REVO Region", "Item", each if Text.Contains([SKU], ":") 
then Text.Trim(Text.AfterDelimiter([SKU], ":")) 
else Text.Trim([SKU])),
    #"Reordered Columns3" = Table.ReorderColumns(#"Added Custom1",{"SKU", "Item", "Forecast Qty", "Regional ID", "YearMonth", "REVO Region", "Product Category"}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Reordered Columns3",{{"Item", type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Changed Type4",{"Item", "Forecast Qty", "Regional ID", "YearMonth", "REVO Region", "Product Category", "SKU", "Sub Location", "Year", "Month", "Lead Group", "Territory"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns1",{"SKU", "Year", "Month"}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Removed Columns1",{{"YearMonth", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type5", each [Forecast Qty] <> 0 and [Forecast Qty] <> null)
in
    #"Filtered Rows";
