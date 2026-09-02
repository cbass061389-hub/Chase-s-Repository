shared #"HIE Shipment" = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Update Meeting Workbook.xlsm"), null, true),
    #"HIE Shipment_Sheet" = Source{[Item="HIE Shipment",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"HIE Shipment_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"STATUS", type text}, {"SUBMITTED", type any}, {"CONFIRMED", type any}, {"RELEASED", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}})
in
    #"Changed Type";
