shared #"New Product Allocations" = let
    Source = Csv.Document(Web.Contents("https://3492685.app.netsuite.com/core/media/media.nl?id=2620991&c=3492685&h=<REDACTED>&_xt=.csv"),[Delimiter=",", Columns=12, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Internal ID", Int64.Type}, {"Name", type text}, {"Description", type text}, {"Base Price", type number}, {"Targeted Launch Date", type date}, {"B2C", Int64.Type}, {"Sponsorships", Int64.Type}, {"EMEA", Int64.Type}, {"Americas", Int64.Type}, {"ASIA", Int64.Type}, {"Tradeshows", Int64.Type}, {"Marketing", Int64.Type}}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Internal ID", "Name", "Description", "Base Price", "Targeted Launch Date"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";
