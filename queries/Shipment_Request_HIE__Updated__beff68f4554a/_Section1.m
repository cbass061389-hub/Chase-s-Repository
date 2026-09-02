section Section1;

shared Open_Order_Report = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600949&c=3492685&h=<REDACTED>&_xt=.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Ship Date", type date}, {"Document Number", type text}, {"Sales Rep", type text}, {"Territory", type text}, {"Company Name", type text}, {"Cust Terms", type text}, {"SO Terms", type text}, {"Status", type text}, {"Inventory Location", type text}, {"PO/Check Number", type text}, {"Item", type text}, {"Item Rate", type number}, {"Amount Discounted", type number}, {"Qty Committed", Int64.Type}, {"Amount Ready To Ship", type number}, {"Back Order Qty", Int64.Type}, {"Back Order Amount", type number}, {"Commit", type text}, {"Amount Unbilled", type number}, {"Balance", type number}, {"Credit Limit", Int64.Type}, {"Credit Limit Notes", type text}, {"Credit Hold", type text}, {"Product Category", type text}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Ship Date", "Territory", "Item"}, {{"BO", each List.Sum([Back Order Qty]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "Days From Today", each Duration.Days([Ship Date]-Date.From(DateTime.LocalNow()))),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"Days From Today", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type2", "dFulfillment Location", each if [Territory] = "B2B China" then "Direct" else if [Territory] = "B2B United States" then "JAX" else if [Territory] = "B2B Asia Pacific" then "Direct" else if [Territory] = "B2B Europe" then "Direct" else if [Territory] = "B2B Canada" then "JAX" else if [Territory] = "B2C United States" then "JAX" else if [Territory] = "B2B Middle East & Africa" then "JAX" else if [Territory] = "B2B Japan" then "Direct" else if [Territory] = "B2B China/Supplier" then "JAX" else if [Territory] = "B2B Latin America" then "JAX" else null),
    #"Added Conditional Column1" = Table.AddColumn(#"Added Conditional Column", "Tradeline", each if [Territory] = "B2B China" then "AP" else if [Territory] = "B2B United States" then "Americas" else if [Territory] = "B2B Asia Pacific" then "AP" else if [Territory] = "B2B Europe" then "EMEA" else if [Territory] = "B2B Canada" then "Americas" else if [Territory] = "B2B Latin America" then "Americas" else if [Territory] = "B2B Middle East & Africa" then "EMEA" else if [Territory] = "B2B Japan" then "AP" else null)
in
    #"Added Conditional Column1";

shared Location_Inventory = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Inventory Location", "Name"}, {{"Available", each List.Sum([On Hand]), type nullable number}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Grouped Rows",{{"Available", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Custom", each if [Inventory Location] = "US B2B" then "JAX" else if [Inventory Location] = "US B2C" then "JAX" else if [Inventory Location] = "INT B2B" then "JAX" else if [Inventory Location] = "INT B2C" then "JAX" else null),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Custom] = "JAX") and ([Available] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom.1", each if Text.Contains([Name], ":") then Text.AfterDelimiter([Name], ":") else [Name]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Inventory Location", "Custom.1", "Name", "Available", "Custom"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Custom.1", "Name"}})
in
    #"Renamed Columns";

shared PO = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Expected Receipt Date", type date}, {"Original PI Date", type date}, {"Legacy PO #", type text}, {"Status", type text}, {"Document Number", type text}, {"Name", type text}, {"Inventory Location", type text}, {"Line ID", Int64.Type}, {"Product Category", type text}, {"Item", type text}, {"Quantity", Int64.Type}, {"Quantity Fulfilled/Received", Int64.Type}, {"Quantity Remaining", Int64.Type}, {"Amount", type number}, {"$ Remaining", type number}, {"Closed", type text}, {"Memo (Main)", type text}, {"Item Note", type text}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Item", "Expected Receipt Date", "Document Number", "Line ID", "Original PI Date"}, {{"Qty", each List.Sum([Quantity Remaining]), type nullable number}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Grouped Rows",{"Item", "Expected Receipt Date", "Qty", "Document Number"}),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Reordered Columns", {{"Document Number", each Text.AfterDelimiter(_, "O"), type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Extracted Text After Delimiter",{"Item", "Expected Receipt Date", "Qty", "Document Number", "Line ID"})
in
    #"Reordered Columns1";

shared #"NetSuite Forecast" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2606291&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=42, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Item", type text}, {"Internal ID", Int64.Type}, {"Product Category", type text}, {"Territory", type text}, {"Sub Location", type text}, {"Lead Group", type text}, {"Forecast QTY 01-2025", type text}, {"Forecast QTY 02-2025", type text}, {"Forecast QTY 03-2025", type text}, {"Forecast QTY 04-2025", type text}, {"Forecast QTY 05-2025", type text}, {"Forecast QTY 06-2025", Int64.Type}, {"Forecast QTY 07-2025", Int64.Type}, {"Forecast QTY 08-2025", Int64.Type}, {"Forecast QTY 09-2025", Int64.Type}, {"Forecast QTY 10-2025", Int64.Type}, {"Forecast QTY 11-2025", Int64.Type}, {"Forecast QTY 12-2025", Int64.Type}, {"Forecast QTY 01-2026", Int64.Type}, {"Forecast QTY 02-2026", Int64.Type}, {"Forecast QTY 03-2026", Int64.Type}, {"Forecast QTY 04-2026", Int64.Type}, {"Forecast QTY 05-2026", Int64.Type}, {"Forecast QTY 06-2026", Int64.Type}, {"Forecast QTY 07-2026", Int64.Type}, {"Forecast QTY 08-2026", Int64.Type}, {"Forecast QTY 09-2026", Int64.Type}, {"Forecast QTY 10-2026", Int64.Type}, {"Forecast QTY 11-2026", Int64.Type}, {"Forecast QTY 12-2026", Int64.Type}, {"Forecast QTY 01-2027", Int64.Type}, {"Forecast QTY 02-2027", Int64.Type}, {"Forecast QTY 03-2027", Int64.Type}, {"Forecast QTY 04-2027", Int64.Type}, {"Forecast QTY 05-2027", Int64.Type}, {"Forecast QTY 06-2027", Int64.Type}, {"Forecast QTY 07-2027", Int64.Type}, {"Forecast QTY 08-2027", Int64.Type}, {"Forecast QTY 09-2027", Int64.Type}, {"Forecast QTY 10-2027", Int64.Type}, {"Forecast QTY 11-2027", Int64.Type}, {"Forecast QTY 12-2027", Int64.Type}}),
    #"Unpivoted Only Selected Columns" = Table.Unpivot(#"Changed Type", {"Forecast QTY 12-2027", "Forecast QTY 11-2027", "Forecast QTY 10-2027", "Forecast QTY 09-2027", "Forecast QTY 08-2027", "Forecast QTY 07-2027", "Forecast QTY 06-2027", "Forecast QTY 05-2027", "Forecast QTY 04-2027", "Forecast QTY 03-2027", "Forecast QTY 02-2027", "Forecast QTY 01-2027", "Forecast QTY 12-2026", "Forecast QTY 11-2026", "Forecast QTY 10-2026", "Forecast QTY 09-2026", "Forecast QTY 08-2026", "Forecast QTY 07-2026", "Forecast QTY 06-2026", "Forecast QTY 05-2026", "Forecast QTY 04-2026", "Forecast QTY 03-2026", "Forecast QTY 02-2026", "Forecast QTY 01-2026", "Forecast QTY 12-2025", "Forecast QTY 11-2025", "Forecast QTY 10-2025", "Forecast QTY 09-2025", "Forecast QTY 08-2025", "Forecast QTY 07-2025", "Forecast QTY 06-2025", "Forecast QTY 05-2025", "Forecast QTY 04-2025", "Forecast QTY 03-2025", "Forecast QTY 02-2025", "Forecast QTY 01-2025"}, "Attribute", "Value"),
    #"Inserted Last Characters" = Table.AddColumn(#"Unpivoted Only Selected Columns", "Last Characters", each Text.End([Attribute], 7), type text),
    #"Added Conditional Column" = Table.AddColumn(#"Inserted Last Characters", "Allocation Group", each if [Territory] = "B2B United States" then "Americas" else if [Territory] = "B2B Latin America" then "Americas" else if [Territory] = "B2B Asia Pacific" then "AP" else if [Territory] = "B2B Japan" then "AP" else if [Territory] = "B2B China" then "AP" else if [Territory] = "B2B Europe" then "EMEA" else if [Territory] = "B2B Middle East & Africa" then "EMEA" else if [Territory] = "B2C Europe" then "EMEA" else if [Lead Group] = "Tradeshows" then "Tradeshows" else if [Lead Group] = "Amazon" then "B2C" else if [Lead Group] = "eComm" then "B2C" else if [Lead Group] = "PBS Events" then "PBS Events" else if [Lead Group] = "Sponsorships" then "Sponsorships" else null),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Conditional Column",{{"Last Characters", type text}, {"Allocation Group", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Last Characters", "Forecast Mnth"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns",{{"Forecast Mnth", type date}, {"Value", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type2", {"Item", "Allocation Group", "Forecast Mnth", "Product Category"}, {{"Forecast Qty", each List.Sum([Value]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows", each ([Forecast Qty] <> null and [Forecast Qty] <> 0)),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered Rows",{"Item", "Allocation Group", "Forecast Mnth", "Forecast Qty", "Product Category"})
in
    #"Reordered Columns";

shared #"Regional Inv Loc" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Inventory Location", type text}, {"Name", type text}, {"Region of Origin", type text}, {"Product Category", type text}, {"Replenishment Method", type text}, {"On Hand", Int64.Type}, {"On Sales Order", Int64.Type}, {"Committed", type text}, {"Available", Int64.Type}, {"Backordered", type text}, {"On Purchase Order", Int64.Type}, {"Low Stock Items", Int64.Type}, {"Not on PO", Int64.Type}, {"Preferred Stock Level", Int64.Type}, {"Location Reorder Point", Int64.Type}, {"Safety Stock Level", Int64.Type}, {"Reorder Multiple", type text}, {"Lead Time", Int64.Type}, {"End Of Life", type text}, {"Discontinued", type text}, {"Location Supply Type", type text}, {"Location Demand Source", type text}, {"Average Cost", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Available", each List.Sum([Available]), type nullable number}}),
    #"Added Conditional Column" = Table.AddColumn(#"Grouped Rows", "Region Inv Loc", each if [Inventory Location] = "HIE" then "HIE" else if [Inventory Location] = "China Office" then "China" else if [Inventory Location] = "Direct Shipments" then "China" else if [Inventory Location] = "Sunray" then "China" else null),
    #"Added Custom" = Table.AddColumn(#"Added Conditional Column", "SKU", each if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else [Name]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Name"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"SKU", "Inventory Location", "Available", "Region Inv Loc"})
in
    #"Reordered Columns";

shared #"New Prod Alloc" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2620991&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=12, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Name", type text}, {"Description", type text}, {"Base Price", type number}, {"Targeted Launch Date", type date}, {"B2C", Int64.Type}, {"Sponsorships", Int64.Type}, {"EMEA", Int64.Type}, {"Americas", Int64.Type}, {"ASIA", Int64.Type}, {"Tradeshows", Int64.Type}, {"Marketing", Int64.Type}}),
    #"Unpivoted Only Selected Columns" = Table.Unpivot(#"Changed Type", {"B2C", "Sponsorships", "EMEA", "Americas", "ASIA", "Tradeshows", "Marketing"}, "Attribute", "Value"),
    #"Filtered Rows" = Table.SelectRows(#"Unpivoted Only Selected Columns", each ([Name] <> "BC8 PRE P3 BLU LL" and [Name] <> "BC8 PRE P3 BLU NW" and [Name] <> "BC8 PRE P3 RED LL" and [Name] <> "BC8 PRE P3 RED NW" and [Name] <> "BCP PRE 10K BLK" and [Name] <> "BCP PRE 10K CHS" and [Name] <> "BCP PRE 10K GLD" and [Name] <> "BCP PRE 10K PUR" and [Name] <> "BCP PRE BK RUSH 2 NW" and [Name] <> "BCP PRE BK RUSH 2 SW" and [Name] <> "BCP PRE BK RUSH OVRD NW" and [Name] <> "BCP PRE BK RUSH OVRD SW" and [Name] <> "BCP PRE BK RUSH VNM NW" and [Name] <> "BCP PRE BK RUSH VNM SW" and [Name] <> "BCP PRE BLA5 1" and [Name] <> "BCP PRE BLA5 2" and [Name] <> "BCP PRE BLA5 3" and [Name] <> "BCP PRE BLA5 4" and [Name] <> "BCP PRE BLA5 5" and [Name] <> "BCP PRE BLA5 CE 1" and [Name] <> "BCP PRE BLA5 CE 2" and [Name] <> "C PRE ROAD 2B4S GRY/BLK H" and [Name] <> "C PRE ROAD 3B5S GRY/BLK H" and [Name] <> "C PRE ROAD 4B8S GRY/BLK S"))
in
    #"Filtered Rows";

shared Sales = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2600950&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=15, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Sum of Quantity", Int64.Type}, {"Date", type date}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Date", "Item", "Territory"}, {{"Quantity", each List.Sum([Sum of Quantity]), type nullable number}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Grouped Rows",{{"Quantity", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Regional ID", each if [Territory] = "B2B United States" then "Americas" else if [Territory] = "B2B Canada" then "Americas" else if [Territory] = "B2B Latin America" then "Americas" else if [Territory] = "B2B Europe" then "EMEA" else if [Territory] = "Beckmann" then "EMEA" else if [Territory] = "B2B Middle East & Africa" then "EMEA" else if [Territory] = "B2B China" then "AP" else if [Territory] = "B2B Japan" then "AP" else if [Territory] = "B2B Asia Pacific" then "AP" else if [Territory] = "B2C United States" then "B2C" else if [Territory] = "Tradeshows" then "Tradeshows" else if [Territory] = "Sponsorships" then "Sponsorships" else null),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Regional ID] <> null))
in
    #"Filtered Rows";

shared #"HIE Shipment" = let
    Source = Excel.Workbook(File.Contents("C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Update Meeting Workbook.xlsm"), null, true),
    #"HIE Shipment_Sheet" = Source{[Item="HIE Shipment",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"HIE Shipment_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"STATUS", type text}, {"SUBMITTED", type any}, {"CONFIRMED", type any}, {"RELEASED", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}})
in
    #"Changed Type";