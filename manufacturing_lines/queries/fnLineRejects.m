// fnLineRejects — one line's Reject Log in the common shape. Load: connection only.
// Layout "TABLE": table tbl_RejectLog: Date, SKU, WO#, Unit ID, Quantity, Description, Serial, Station.
// Layout "REVO":  sheet "Reject Log" with headers Date, WO, SKU, Cart, RejectQty, Reason, User.
//   The REVO macro (REVO_Ops.LogReject) writes the SKU into the "WO" column and the WO into the "SKU" column,
//   so the two are un-swapped here. Verify against the live REVO file and delete the swap if it was fixed.
(line as record) as table =>
let
    wb       = Excel.Workbook(File.Contents(line[Workbook path]), null, true),
    raw      = if line[Layout] = "REVO"
               then Table.PromoteHeaders(wb{[Item = "Reject Log", Kind = "Sheet"]}[Data], [PromoteAllScalars = true])
               else wb{[Item = "tbl_RejectLog", Kind = "Table"]}[Data],
    step1    = if line[Layout] = "REVO" then Table.RenameColumns(raw,   {{"SKU", "WO_tmp"}}, MissingField.Ignore) else raw,
    step2    = if line[Layout] = "REVO" then Table.RenameColumns(step1, {{"WO", "SKU"}}, MissingField.Ignore) else step1,
    renamed  = if line[Layout] = "REVO"
               then Table.RenameColumns(step2, {{"WO_tmp", "WO"}, {"Cart", "Unit ID"}, {"RejectQty", "Quantity"}, {"Reason", "Description"}}, MissingField.Ignore)
               else Table.RenameColumns(step2, {{"WO#", "WO"}}, MissingField.Ignore),
    selected = Table.SelectColumns(renamed, {"Date", "SKU", "WO", "Unit ID", "Quantity", "Description", "Serial", "Station"}, MissingField.UseNull),
    typed    = Table.TransformColumnTypes(selected, {{"Date", type date}, {"Quantity", type number}, {"SKU", type text}, {"WO", type text},
                                                     {"Unit ID", type text}, {"Description", type text}, {"Serial", type text}, {"Station", type text}}, "en-US"),
    nonEmpty = Table.SelectRows(typed, each [Date] <> null and [SKU] <> null and [SKU] <> ""),
    withLine = Table.AddColumn(nonEmpty, "Line", each line[Line], type text),
    ordered  = Table.ReorderColumns(withLine, {"Line", "Date", "SKU", "WO", "Unit ID", "Quantity", "Description", "Serial", "Station"})
in
    ordered
