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
