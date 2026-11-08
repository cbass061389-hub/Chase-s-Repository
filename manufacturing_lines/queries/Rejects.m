// Rejects — every enabled line's Reject Log, one shape. Load to: Data_Rejects!A1.
let
    perLine  = Table.AddColumn(Lines, "Data", each fnLineRejects(_)),
    combined = Table.Combine(perLine[Data]),
    typed    = Table.TransformColumnTypes(combined, {{"Line", type text}, {"Date", type date}, {"SKU", type text}, {"WO", type text},
                                                     {"Unit ID", type text}, {"Quantity", type number}, {"Description", type text},
                                                     {"Serial", type text}, {"Station", type text}}),
    sorted   = Table.Sort(typed, {{"Date", Order.Descending}, {"Line", Order.Ascending}})
in
    sorted
