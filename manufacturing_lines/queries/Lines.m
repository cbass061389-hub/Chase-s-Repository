// Lines — the three line workbooks, from the 'Lines' table on Config. Load: connection only.
let
    Source  = Excel.CurrentWorkbook(){[Name = "Lines"]}[Content],
    Typed   = Table.TransformColumnTypes(Source, {{"Line", type text}, {"Workbook path", type text}, {"Enabled", type logical}, {"Layout", type text}}),
    Enabled = Table.SelectRows(Typed, each [Enabled] = true)
in
    Enabled
