shared #"Purchase Price" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2613629&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=36, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Type", type text}, {"UPC Code", type text}, {"Name", type text}, {"Item Image", type text}, {"Display Name", type text}, {"Item Parent Sub", type text}, {"Product Category", type text}, {"Product Category (no hierarchy)", type text}, {"Replenishment Method", type text}, {"Purchase Description", type text}, {"Sponsorship Value", type text}, {"Inactive", type text}, {"End Of Life", type text}, {"Item Country Of Origin", type text}, {"Sales Description", type text}, {"Vendor", type text}, {"Vendor Name", type text}, {"Drawing Revision #", type text}, {"HTS Code", type text}, {"HTS Code Description", type text}, {"Weight", type number}, {"Weight Units", type text}, {"Sync to Predator Group LSX", type text}, {"Print On Price List", type text}, {"Track in NSPB", type text}, {"Internal ID", Int64.Type}, {"Joint Type", type text}, {"Last Purchase Price", type number}, {"Vendor Price", type number}, {"Average Cost", type number}, {"Core Item", type text}, {"Targeted Launch Date", type text}, {"End Of Product Life Cycle", type text}, {"Critical Item for Production", type text}, {"Min Order Qty", type text}, {"Reorder Point", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Item Parent Sub", "Product Category (no hierarchy)"}, {{"Average Cost", each List.Max([Average Cost]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "SKU", each if Text.Contains([Name], ":") 
then Text.Trim(Text.AfterDelimiter([Name], ":")) 
else Text.Trim([Name])),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Name"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "Item Parent Sub", "Product Category (no hierarchy)", "Average Cost"})
in
    #"Reordered Columns";
