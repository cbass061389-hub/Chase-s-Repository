shared PO = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Expected Receipt Date", type date}, {"Original PI Date", type date}, {"Legacy PO #", type text}, {"Status", type text}, {"Document Number", type text}, {"Name", type text}, {"Inventory Location", type text}, {"Line ID", Int64.Type}, {"Product Category", type text}, {"Item", type text}, {"Quantity", Int64.Type}, {"Quantity Fulfilled/Received", Int64.Type}, {"Quantity Remaining", Int64.Type}, {"Amount", type number}, {"$ Remaining", type number}, {"Closed", type text}, {"Memo (Main)", type text}, {"Item Note", type text}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Item", "Expected Receipt Date", "Document Number", "Line ID", "Original PI Date"}, {{"Qty", each List.Sum([Quantity Remaining]), type nullable number}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Grouped Rows",{"Item", "Expected Receipt Date", "Qty", "Document Number"}),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Reordered Columns", {{"Document Number", each Text.AfterDelimiter(_, "O"), type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Extracted Text After Delimiter",{"Item", "Expected Receipt Date", "Qty", "Document Number", "Line ID"})
in
    #"Reordered Columns1";
