shared #"Items On PO" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Expected Receipt Date", type date}, {"Original PI Date", type date}, {"Legacy PO #", type text}, {"Status", type text}, {"Document Number", type text}, {"Name", type text}, {"Inventory Location", type text}, {"Line ID", Int64.Type}, {"Product Category", type text}, {"Item", type text}, {"Quantity", Int64.Type}, {"Quantity Fulfilled/Received", Int64.Type}, {"Quantity Remaining", Int64.Type}, {"Amount", type number}, {"$ Remaining", type number}, {"Closed", type text}, {"Memo (Main)", type text}, {"Item Note", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Expected Receipt Date", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows",{{"Item", "SKU"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Item", each if Text.Contains([SKU], ":") 
then Text.Trim(Text.AfterDelimiter([SKU], ":")) 
else Text.Trim([SKU])),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Date", "Expected Receipt Date", "Original PI Date", "Legacy PO #", "Status", "Document Number", "Name", "Inventory Location", "Line ID", "Product Category", "SKU", "Item", "Quantity", "Quantity Fulfilled/Received", "Quantity Remaining", "Amount", "$ Remaining", "Closed", "Memo (Main)", "Item Note"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"SKU"})
in
    #"Removed Columns";
