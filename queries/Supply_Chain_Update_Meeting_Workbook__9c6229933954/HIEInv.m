shared HIEInv = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE .xlsm"), null, true),
    HIEInv_Sheet = Source{[Item="HIEInv",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(HIEInv_Sheet, [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers",{"PO"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"SKU", type text}, {"Qty", Int64.Type}, {"Source", type text},{"Remark", type any}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"SKU", "Source"}, {{"Ttl Qty", each List.Sum([Qty]), type nullable number}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"SKU"}, {{"HIE Inv", each List.Sum([Ttl Qty]), type nullable number}})
in
    #"Grouped Rows1";
