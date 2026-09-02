section Section1;

shared #"Current Inventory" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Ttl Available", each List.Sum([On Hand]), type nullable number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Name", "Item"}, {"Ttl Available", "InvQty"}, {"Inventory Location", "InvLoc"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([InvQty] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Regional ID", each if [InvLoc] = "INT Tradeshows" then "Tradeshow" else if [InvLoc] = "INT B2C" then "B2C" else if [InvLoc] = "INT B2B" then "Americas" else if [InvLoc] = "China Office" then "China" else if [InvLoc] = "US B2C" then "B2C" else if [InvLoc] = "US B2B" then "Americas" else if [InvLoc] = "Sunray" then "China" else if [InvLoc] = "HIE" then "China" else if [InvLoc] = "Jax Tradeshows" then "Tradeshow" else if [InvLoc] = "Direct Shipments" then "China" else null),
    #"Added Conditional Column" = Table.AddColumn(#"Added Custom", "Custom", each if [Regional ID] = "B2C" then "B2C" else if [Regional ID] = "China" then "AP" else if [Regional ID] = "Americas" then "Americas" else if [Regional ID] = "EMEA" then "EMEA" else null),
    #"Renamed Columns1" = Table.RenameColumns(#"Added Conditional Column",{{"Custom", "REVO Region"}}),
    #"Filtered Rows1" = Table.SelectRows(#"Renamed Columns1", each true),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows1", "SKU", each if Text.Contains([Item], ":") 
then Text.Trim(Text.AfterDelimiter([Item], ":")) 
else Text.Trim([Item])),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{"Item"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "InvLoc", "InvQty", "Regional ID", "REVO Region"})
in
    #"Reordered Columns";

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

shared #"Items On PO" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Expected Receipt Date", type date}, {"Original PI Date", type date}, {"Legacy PO #", type text}, {"Status", type text}, {"Document Number", type text}, {"Name", type text}, {"Inventory Location", type text}, {"Line ID", Int64.Type}, {"Product Category", type text}, {"Item", type text}, {"Quantity", Int64.Type}, {"Quantity Fulfilled/Received", Int64.Type}, {"Quantity Remaining", Int64.Type}, {"Amount", type number}, {"$ Remaining", type number}, {"Closed", type text}, {"Memo (Main)", type text}, {"Item Note", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Expected Receipt Date", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows",{{"Item", "SKU"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Item", each if Text.Contains([SKU], ":") 
then Text.Trim(Text.AfterDelimiter([SKU], ":")) 
else Text.Trim([SKU])),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Date", "Expected Receipt Date", "Original PI Date", "Legacy PO #", "Status", "Document Number", "Name", "Inventory Location", "Line ID", "Product Category", "SKU", "Item", "Quantity", "Quantity Fulfilled/Received", "Quantity Remaining", "Amount", "$ Remaining", "Closed", "Memo (Main)", "Item Note"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"SKU"})
in
    #"Removed Columns";

shared #"New Product Allocations" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2620991&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=12, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Name", type text}, {"Description", type text}, {"Base Price", type number}, {"Targeted Launch Date", type date}, {"B2C", Int64.Type}, {"Sponsorships", Int64.Type}, {"EMEA", Int64.Type}, {"Americas", Int64.Type}, {"ASIA", Int64.Type}, {"Tradeshows", Int64.Type}, {"Marketing", Int64.Type}}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Internal ID", "Name", "Description", "Base Price", "Targeted Launch Date"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

shared #"Purchase Price" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2613629&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=36, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Type", type text}, {"UPC Code", type text}, {"Name", type text}, {"Item Image", type text}, {"Display Name", type text}, {"Item Parent Sub", type text}, {"Product Category", type text}, {"Product Category (no hierarchy)", type text}, {"Replenishment Method", type text}, {"Purchase Description", type text}, {"Sponsorship Value", type text}, {"Inactive", type text}, {"End Of Life", type text}, {"Item Country Of Origin", type text}, {"Sales Description", type text}, {"Vendor", type text}, {"Vendor Name", type text}, {"Drawing Revision #", type text}, {"HTS Code", type text}, {"HTS Code Description", type text}, {"Weight", type number}, {"Weight Units", type text}, {"Sync to Predator Group LSX", type text}, {"Print On Price List", type text}, {"Track in NSPB", type text}, {"Internal ID", Int64.Type}, {"Joint Type", type text}, {"Last Purchase Price", type number}, {"Vendor Price", type number}, {"Average Cost", type number}, {"Core Item", type text}, {"Targeted Launch Date", type text}, {"End Of Product Life Cycle", type text}, {"Critical Item for Production", type text}, {"Min Order Qty", type text}, {"Reorder Point", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Item Parent Sub", "Product Category (no hierarchy)"}, {{"Average Cost", each List.Max([Average Cost]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "SKU", each if Text.Contains([Name], ":") 
then Text.Trim(Text.AfterDelimiter([Name], ":")) 
else Text.Trim([Name])),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Name"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "Item Parent Sub", "Product Category (no hierarchy)", "Average Cost"})
in
    #"Reordered Columns";

shared #"Forecast Qty" = let
    Source = Csv.Document(
        Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2606291&c=3492685&h=<REDACTED>&_xt=.csv"),
        [Delimiter=",", Columns=42, Encoding=65001, QuoteStyle=QuoteStyle.None]
    ),

    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    #"Changed Type" = Table.TransformColumnTypes(
        #"Promoted Headers",
        {
            {"Item", type text},
            {"Internal ID", Int64.Type},
            {"Product Category", type text},
            {"Territory", type text},
            {"Sub Location", type text},
            {"Lead Group", type text},
            {"Forecast QTY 01-2025", type text},
            {"Forecast QTY 02-2025", type text},
            {"Forecast QTY 03-2025", type text},
            {"Forecast QTY 04-2025", type text},
            {"Forecast QTY 05-2025", type text},
            {"Forecast QTY 06-2025", Int64.Type},
            {"Forecast QTY 07-2025", Int64.Type},
            {"Forecast QTY 08-2025", Int64.Type},
            {"Forecast QTY 09-2025", Int64.Type},
            {"Forecast QTY 10-2025", Int64.Type},
            {"Forecast QTY 11-2025", Int64.Type},
            {"Forecast QTY 12-2025", Int64.Type},
            {"Forecast QTY 01-2026", Int64.Type},
            {"Forecast QTY 02-2026", Int64.Type},
            {"Forecast QTY 03-2026", Int64.Type},
            {"Forecast QTY 04-2026", Int64.Type},
            {"Forecast QTY 05-2026", Int64.Type},
            {"Forecast QTY 06-2026", Int64.Type},
            {"Forecast QTY 07-2026", Int64.Type},
            {"Forecast QTY 08-2026", Int64.Type},
            {"Forecast QTY 09-2026", Int64.Type},
            {"Forecast QTY 10-2026", Int64.Type},
            {"Forecast QTY 11-2026", Int64.Type},
            {"Forecast QTY 12-2026", Int64.Type},
            {"Forecast QTY 01-2027", Int64.Type},
            {"Forecast QTY 02-2027", Int64.Type},
            {"Forecast QTY 03-2027", Int64.Type},
            {"Forecast QTY 04-2027", Int64.Type},
            {"Forecast QTY 05-2027", Int64.Type},
            {"Forecast QTY 06-2027", Int64.Type},
            {"Forecast QTY 07-2027", Int64.Type},
            {"Forecast QTY 08-2027", Int64.Type},
            {"Forecast QTY 09-2027", Int64.Type},
            {"Forecast QTY 10-2027", Int64.Type},
            {"Forecast QTY 11-2027", Int64.Type},
            {"Forecast QTY 12-2027", Int64.Type}
        }
    ),

    // Turn monthly columns into rows
    #"Unpivoted Other Columns" =
        Table.UnpivotOtherColumns(
            #"Changed Type",
            {"Lead Group", "Sub Location", "Territory", "Product Category", "Internal ID", "Item"},
            "Attribute",
            "Value"
        ),
    // Keep only the "MM-YYYY" piece and convert to date
    #"Extracted Last Characters" =
        Table.TransformColumns(
            #"Unpivoted Other Columns",
            {{"Attribute", each Text.End(_, 7), type text}}
        ),

    #"Changed Type1" =
        Table.TransformColumnTypes(#"Extracted Last Characters", {{"Attribute", type date}}),

    #"Renamed Columns" =
        Table.RenameColumns(#"Changed Type1", {{"Value", "Forecast Qty"}}),

    #"Changed Type2" =
        Table.TransformColumnTypes(#"Renamed Columns", {{"Forecast Qty", Int64.Type}}),

    // Build SKU from Item (same as you had)
    #"Added Custom" =
        Table.AddColumn(
            #"Changed Type2",
            "SKU",
            each if Text.Contains([Item], ":") then Text.AfterDelimiter([Item], ":") else [Item]
        ),

    #"Reordered Columns" =
        Table.ReorderColumns(
            #"Added Custom",
            {"SKU", "Item", "Internal ID", "Product Category", "Territory", "Sub Location", "Lead Group", "Attribute", "Forecast Qty"}
        ),

    #"Removed Columns" =
        Table.RemoveColumns(#"Reordered Columns", {"Item"}),

    // Extract Year/Month from the Attribute date
    #"Added Year" =
        Table.AddColumn(#"Removed Columns", "Year", each Date.Year([Attribute]), Int64.Type),

    #"Added Month" =
        Table.AddColumn(#"Added Year", "Month", each Date.Month([Attribute]), Int64.Type),

    // Aggregate to monthly SKU × Sub Location
    #"Grouped Rows1" =
        Table.Group(#"Added Month", {"SKU", "Sub Location", "Year", "Month", "Product Category", "Lead Group", "Territory"}, {{"Forecast Qty", each List.Sum([Forecast Qty]), type nullable number}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Grouped Rows1",{"SKU", "Sub Location", "Year", "Month", "Forecast Qty", "Product Category"}),

    // Map Sub Location → Regional ID (mirroring Forecast query logic)
    #"Added Regional ID" =
        Table.AddColumn(#"Reordered Columns2", "Regional ID", each if [Lead Group] = "Amazon" then "Amazon" else if [Lead Group] = "PBS Events" then "PBS Events" else if [Lead Group] = "eComm" then "eComm"
                else if [Lead Group] = "Sponsorships Representing" or [Lead Group] = "Sponsorships Product Compensation" then "Sponsorship"
                else if [Lead Group] = "Tradeshows" then "Tradeshows"
                else if [Lead Group] = "Beckmann" then "EMEA"
                else if [Sub Location] = "Americas" then "Americas"
                else if [Sub Location] = "Asia" then "Asia"
                else if [Sub Location] = "EMEA" then "EMEA"
                else null),

    #"Changed Type3" =
        Table.TransformColumnTypes(
            #"Added Regional ID",
            {{"Regional ID", type text}, {"Forecast Qty", Int64.Type}}
        ),

    // Build YearMonth like in the Forecast query
    #"Added YearMonth" =
        Table.AddColumn(#"Changed Type3", "YearMonth", each Text.From([Year])& Text.PadStart(Text.From([Month]), 2, "0")),

    // REVO Region mirroring the template
    #"Added REVO Region" =
        Table.AddColumn(
            #"Added YearMonth",
            "REVO Region",
            each
                if [Regional ID] = "Americas" then "Americas"
                else if [Regional ID] = "EMEA" then "EMEA"
                else if [Regional ID] = "AP" then "AP"
                else if [Regional ID] = "B2C" then "B2C"
                else null
        ),

    #"Added Custom1" = Table.AddColumn(#"Added REVO Region", "Item", each if Text.Contains([SKU], ":") 
then Text.Trim(Text.AfterDelimiter([SKU], ":")) 
else Text.Trim([SKU])),
    #"Reordered Columns3" = Table.ReorderColumns(#"Added Custom1",{"SKU", "Item", "Forecast Qty", "Regional ID", "YearMonth", "REVO Region", "Product Category"}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Reordered Columns3",{{"Item", type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Changed Type4",{"Item", "Forecast Qty", "Regional ID", "YearMonth", "REVO Region", "Product Category", "SKU", "Sub Location", "Year", "Month", "Lead Group", "Territory"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns1",{"SKU", "Year", "Month"}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Removed Columns1",{{"YearMonth", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type5", each [Forecast Qty] <> 0 and [Forecast Qty] <> null)
in
    #"Filtered Rows";

shared HIEInv = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE .xlsm"), null, true),
    HIEInv_Sheet = Source{[Item="HIEInv",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(HIEInv_Sheet, [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers",{"PO"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"SKU", type text}, {"Qty", Int64.Type}, {"Source", type text},{"Remark", type any}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"SKU", "Source"}, {{"Ttl Qty", each List.Sum([Qty]), type nullable number}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"SKU"}, {{"HIE Inv", each List.Sum([Ttl Qty]), type nullable number}})
in
    #"Grouped Rows1";

shared HIEProdFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    HIEProdFlat_Sheet = Source{[Item="HIEProdFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(HIEProdFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type", {"Date"})
in
    #"Removed Errors";

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

shared JingdianProdFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    JingdianProdFlat_Sheet = Source{[Item="JingdianProdFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(JingdianProdFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"PO", Int64.Type}, {"SKU", type text}, {"Chinese Item", type text}, {"Date", type date}, {"Qty", Int64.Type}, {"Status", type text}, {"Source Sheet", type text}})
in
    #"Changed Type";

shared PoisonInvFlat = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - Updated.xlsm"), null, true),
    PoisonInvFlat_Sheet = Source{[Item="PoisonInvFlat",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(PoisonInvFlat_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"SKU", type text}, {"Qty", Int64.Type}, {"PO", type any}, {"Unit Price", Int64.Type}, {"Amount", Int64.Type}, {"Source Sheet", type text}})
in
    #"Changed Type";