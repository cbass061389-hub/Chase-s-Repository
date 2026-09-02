shared Location_Inventory = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Inventory Location", "Name"}, {{"Available", each List.Sum([On Hand]), type nullable number}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Grouped Rows",{{"Available", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Custom", each if [Inventory Location] = "US B2B" then "JAX" else if [Inventory Location] = "US B2C" then "JAX" else if [Inventory Location] = "INT B2B" then "JAX" else if [Inventory Location] = "INT B2C" then "JAX" else null),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Custom] = "JAX") and ([Available] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom.1", each if Text.Contains([Name], ":") then Text.AfterDelimiter([Name], ":") else [Name]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Inventory Location", "Custom.1", "Name", "Available", "Custom"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Custom.1", "Name"}})
in
    #"Renamed Columns";
