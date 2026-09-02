shared PoisonInvFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    PoisonInvFlat_Sheet = Source{[Item="PoisonInvFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(PoisonInvFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"SKU", type text}, {"Qty", Int64.Type}, {"PO", type any}, {"Unit Price", Int64.Type}, {"Amount", Int64.Type}, {"Source Sheet", type text}})
in
    #"Changed Type";
