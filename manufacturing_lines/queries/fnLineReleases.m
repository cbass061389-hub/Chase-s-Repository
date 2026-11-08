// fnLineReleases — one line's Release Log in the common shape. Load: connection only.
// Layout "TABLE": table tbl_ReleaseLog (TABLE and WOOD workbooks; real dates).
// Layout "REVO":  sheet "Release Log" with headers Date, SKU, Cart, ReleasedQty, WO, Rework, B Grade, Reject, Verified (text dates m/d/yyyy).
// A missing file or sheet fails the refresh on purpose: a line that silently reads as zero is worse than an error.
(line as record) as table =>
let
    wb       = Excel.Workbook(File.Contents(line[Workbook path]), null, true),
    raw      = if line[Layout] = "REVO"
               then Table.PromoteHeaders(wb{[Item = "Release Log", Kind = "Sheet"]}[Data], [PromoteAllScalars = true])
               else wb{[Item = "tbl_ReleaseLog", Kind = "Table"]}[Data],
    renamed  = if line[Layout] = "REVO"
               then Table.RenameColumns(raw, {{"Cart", "Unit ID"}, {"ReleasedQty", "Qty Released"}, {"WO", "Work Order"},
                                             {"Rework", "Qty Rework"}, {"B Grade", "Qty B Grade"}, {"Reject", "Qty Reject"}}, MissingField.Ignore)
               else raw,
    selected = Table.SelectColumns(renamed, {"Date", "SKU", "Unit ID", "Qty Released", "Qty Rework", "Qty B Grade", "Qty Reject", "Serial", "Work Order"}, MissingField.UseNull),
    typed    = Table.TransformColumnTypes(selected, {{"Date", type date}, {"Qty Released", type number}, {"Qty Rework", type number},
                                                     {"Qty B Grade", type number}, {"Qty Reject", type number}, {"SKU", type text},
                                                     {"Unit ID", type text}, {"Serial", type text}, {"Work Order", type text}}, "en-US"),
    nonEmpty = Table.SelectRows(typed, each [Date] <> null and [SKU] <> null and [SKU] <> ""),
    withLine = Table.AddColumn(nonEmpty, "Line", each line[Line], type text),
    ordered  = Table.ReorderColumns(withLine, {"Line", "Date", "SKU", "Unit ID", "Qty Released", "Qty Rework", "Qty B Grade", "Qty Reject", "Serial", "Work Order"})
in
    ordered
