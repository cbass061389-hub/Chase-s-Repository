// fnLinePlan — one line's Production Plan grid unpivoted to Week Start x SKU. Load: connection only.
// All three workbooks share the grid: column A = SKU, row 2 = week-start dates, SKUs from row 5 (rows 3-4 are capacity and totals).
(line as record) as table =>
let
    wb        = Excel.Workbook(File.Contents(line[Workbook path]), null, true),
    rows      = wb{[Item = "Production Plan", Kind = "Sheet"]}[Data],
    colNames  = Table.ColumnNames(rows),
    weekCols  = List.Skip(colNames, 1),
    headerRow = rows{1},
    dateMap   = Table.FromColumns({weekCols, List.Transform(weekCols, each Record.Field(headerRow, _))}, {"Attribute", "Week Start"}),
    dateMapOk = Table.SelectRows(Table.TransformColumnTypes(dateMap, {{"Week Start", type date}}), each [Week Start] <> null),
    dataRows  = Table.Skip(rows, 4),
    renamed   = Table.RenameColumns(dataRows, {{List.First(colNames), "SKU"}}),
    unpivoted = Table.UnpivotOtherColumns(renamed, {"SKU"}, "Attribute", "Planned Qty"),
    joined    = Table.NestedJoin(unpivoted, {"Attribute"}, dateMapOk, {"Attribute"}, "d", JoinKind.Inner),
    expanded  = Table.ExpandTableColumn(joined, "d", {"Week Start"}),
    typed     = Table.TransformColumnTypes(expanded, {{"SKU", type text}, {"Week Start", type date}, {"Planned Qty", type number}}),
    kept      = Table.SelectRows(typed, each [SKU] <> null and [SKU] <> "" and [Planned Qty] <> null and [Planned Qty] <> 0),
    withLine  = Table.AddColumn(kept, "Line", each line[Line], type text),
    ordered   = Table.SelectColumns(withLine, {"Line", "Week Start", "SKU", "Planned Qty"})
in
    ordered
