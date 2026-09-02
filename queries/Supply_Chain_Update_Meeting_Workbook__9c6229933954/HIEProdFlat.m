shared HIEProdFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    HIEProdFlat_Sheet = Source{[Item="HIEProdFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(HIEProdFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type", {"Date"})
in
    #"Removed Errors";
