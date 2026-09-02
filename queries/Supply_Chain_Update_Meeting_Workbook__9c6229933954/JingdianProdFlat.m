shared JingdianProdFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    JingdianProdFlat_Sheet = Source{[Item="JingdianProdFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(JingdianProdFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"PO", Int64.Type}, {"SKU", type text}, {"Chinese Item", type text}, {"Date", type date}, {"Qty", Int64.Type}, {"Status", type text}, {"Source Sheet", type text}})
in
    #"Changed Type";
