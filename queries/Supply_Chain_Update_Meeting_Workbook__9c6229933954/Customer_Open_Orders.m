shared #"Customer Open Orders" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600949&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=25, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Ship Date", type date}, {"Document Number", type text}, {"Sales Rep", type text}, {"Territory", type text}, {"Company Name", type text}, {"Cust Terms", type text}, {"SO Terms", type text}, {"Status", type text}, {"Inventory Location", type text}, {"PO/Check Number", type text}, {"Item", type text}, {"Item Rate", type number}, {"Amount Discounted", type number}, {"Qty Committed", Int64.Type}, {"Amount Ready To Ship", type number}, {"Back Order Qty", Int64.Type}, {"Back Order Amount", type number}, {"Commit", type text}, {"Amount Unbilled", type number}, {"Balance", type number}, {"Credit Limit", Int64.Type}, {"Credit Limit Notes", type text}, {"Credit Hold", type text}, {"Product Category", type text}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Regional ID", each if [Territory] = "B2B Asia Pacific" then "AP" 
else if [Territory] = "B2B Canada" then "Americas"
else if [Territory] = "B2B China" then "AP"
else if [Territory] = "B2B Europe" then "EMEA"
else if [Territory] = "B2B Japan" then "AP"
else if [Territory] = "B2B Latin America" then "Americas"
else if [Territory] = "B2B Middle East & Africa" then "EMEA"
else if [Territory] = "B2B United States" then "Americas"
else if [Territory] = "B2C Europe" then "B2C"
else if [Territory] = "B2C International" then "B2C"
else if [Territory] = "B2C United States" then "B2C"
else null),
    #"Grouped Rows" = Table.Group(#"Added Custom", {"Item", "Ship Date", "Regional ID", "Document Number", "Company Name", "Commit", "Sales Rep", "Balance", "Credit Limit", "Credit Hold", "Status", "Item Rate"}, {{"TtlBO", each List.Sum([Back Order Qty]), type nullable number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{"Item", "Ship Date", "Regional ID", "TtlBO", "Document Number"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns1",{{"TtlBO", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Item", "SKU"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Item", each if Text.Contains([SKU], ":") 
then Text.Trim(Text.AfterDelimiter([SKU], ":")) 
else Text.Trim([SKU])),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{"SKU"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Item", "Ship Date", "Regional ID", "TtlBO"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Reordered Columns",{{"Item", type text}})
in
    #"Changed Type2";
