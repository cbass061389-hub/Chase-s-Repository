// RejectPareto — reject qty by Line x Reason over the last 13 weeks, ranked within line. Load to: Data_Pareto!A1.
// Reason = Description up to " — " (the QC forms write "reason — notes"); blanks become "(no reason recorded)".
let
    since    = Date.AddWeeks(Date.From(DateTime.LocalNow()), -13),
    recent   = Table.SelectRows(Rejects, each [Date] <> null and [Date] >= since),
    reason   = Table.AddColumn(recent, "Reason", each
                   let d = if [Description] = null then "" else Text.Trim([Description]),
                       r = if Text.Contains(d, " — ") then Text.BeforeDelimiter(d, " — ") else d
                   in  if r = "" then "(no reason recorded)" else r, type text),
    grouped  = Table.Group(reason, {"Line", "Reason"}, {{"Qty", each List.Sum([Quantity]), type number}}),
    sorted   = Table.Sort(grouped, {{"Line", Order.Ascending}, {"Qty", Order.Descending}}),
    ranked   = Table.Combine(Table.Group(sorted, {"Line"}, {{"t", each Table.AddIndexColumn(_, "Rank", 1, 1), type table}})[t]),
    keyed    = Table.AddColumn(ranked, "Key", each [Line] & "|" & Text.From([Rank]), type text),
    typed    = Table.TransformColumnTypes(keyed, {{"Line", type text}, {"Reason", type text}, {"Qty", type number}, {"Rank", Int64.Type}, {"Key", type text}})
in
    typed
