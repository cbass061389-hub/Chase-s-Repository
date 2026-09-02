shared ValidatedSales = let
    Source = Excel.Workbook(Web.Contents("https://predatorgroup.sharepoint.com/Share%20All%20Files/Forecast/Reference/Actuals%20vs%20Forecast%20Archive/Sales%20Actuals%20+%20Forecast%20for%20Actuals%20vs%20Forecast.xlsx"), null, true),
    Custom1 = Source{[Item="SKU Level Actuals",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Custom1, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Location_lvl_1", type text}, {"Location_lvl_2", type text}, {"Location_lvl_3", type text}, {"Product_Line", type text}, {"Product", type text}, {"Product_Desc", type text}, {"Year", Int64.Type}, {"YYYYMM", Int64.Type}, {"Name", type text}, {"Sum of Sales", type number}, {"Sum of Sum of Quantity", Int64.Type}, {"Sum of GP", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Location_lvl_2", "YYYYMM", "Name", "Year"}, {{"Units", each List.Sum([Sum of Sum of Quantity]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows", each ([Year] = 2025 or [Year] = 2026)),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered Rows",{"Location_lvl_2", "YYYYMM", "Name", "Units", "Year"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns", "SKU", each if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else [Name]),
    #"Trimmed Text" = Table.TransformColumns(#"Added Custom",{{"SKU", Text.Trim, type text}}),
    #"Filtered Rows1" = Table.SelectRows(#"Trimmed Text", each ([Location_lvl_2] <> "Cross Subsidy" and [Location_lvl_2] <> "Elimination" and [Location_lvl_2] <> "Non-Inventory"))
in
    #"Filtered Rows1";
