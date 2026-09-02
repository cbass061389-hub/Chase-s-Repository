shared #"Current Inventory" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Ttl Available", each List.Sum([On Hand]), type nullable number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Name", "Item"}, {"Ttl Available", "InvQty"}, {"Inventory Location", "InvLoc"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([InvQty] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Regional ID", each if [InvLoc] = "INT Tradeshows" then "Tradeshow" else if [InvLoc] = "INT B2C" then "B2C" else if [InvLoc] = "INT B2B" then "Americas" else if [InvLoc] = "China Office" then "China" else if [InvLoc] = "US B2C" then "B2C" else if [InvLoc] = "US B2B" then "Americas" else if [InvLoc] = "Sunray" then "China" else if [InvLoc] = "HIE" then "China" else if [InvLoc] = "Jax Tradeshows" then "Tradeshow" else if [InvLoc] = "Direct Shipments" then "China" else null),
    #"Added Conditional Column" = Table.AddColumn(#"Added Custom", "Custom", each if [Regional ID] = "B2C" then "B2C" else if [Regional ID] = "China" then "AP" else if [Regional ID] = "Americas" then "Americas" else if [Regional ID] = "EMEA" then "EMEA" else null),
    #"Renamed Columns1" = Table.RenameColumns(#"Added Conditional Column",{{"Custom", "REVO Region"}}),
    #"Filtered Rows1" = Table.SelectRows(#"Renamed Columns1", each true),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows1", "SKU", each if Text.Contains([Item], ":") 
then Text.Trim(Text.AfterDelimiter([Item], ":")) 
else Text.Trim([Item])),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{"Item"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "InvLoc", "InvQty", "Regional ID", "REVO Region"})
in
    #"Reordered Columns";
