// Releases — every enabled line's Release Log, one shape. Load to: Data_Releases!A1.
let
    perLine  = Table.AddColumn(Lines, "Data", each fnLineReleases(_)),
    combined = Table.Combine(perLine[Data]),
    typed    = Table.TransformColumnTypes(combined, {{"Line", type text}, {"Date", type date}, {"SKU", type text}, {"Unit ID", type text},
                                                     {"Qty Released", type number}, {"Qty Rework", type number}, {"Qty B Grade", type number},
                                                     {"Qty Reject", type number}, {"Serial", type text}, {"Work Order", type text}}),
    sorted   = Table.Sort(typed, {{"Date", Order.Descending}, {"Line", Order.Ascending}})
in
    sorted
