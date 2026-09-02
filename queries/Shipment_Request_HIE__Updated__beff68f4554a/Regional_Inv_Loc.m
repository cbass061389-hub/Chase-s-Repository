shared #"Regional Inv Loc" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Available", each List.Sum([Available]), type nullable number}}),
    #"Added Conditional Column" = Table.AddColumn(#"Grouped Rows", "Region Inv Loc", each if [Inventory Location] = "HIE" then "HIE" else if [Inventory Location] = "China Office" then "China" else if [Inventory Location] = "Direct Shipments" then "China" else if [Inventory Location] = "Sunray" then "China" else null),
    #"Added Custom" = Table.AddColumn(#"Added Conditional Column", "SKU", each if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else [Name]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Name"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "Inventory Location", "Available", "Region Inv Loc"})
in
    #"Reordered Columns";
