// Plan — every enabled line's Production Plan, unpivoted. Load to: Data_Plan!A1.
let
    perLine  = Table.AddColumn(Lines, "Data", each fnLinePlan(_)),
    combined = Table.Combine(perLine[Data]),
    typed    = Table.TransformColumnTypes(combined, {{"Line", type text}, {"Week Start", type date}, {"SKU", type text}, {"Planned Qty", type number}}),
    sorted   = Table.Sort(typed, {{"Week Start", Order.Ascending}, {"Line", Order.Ascending}})
in
    sorted
