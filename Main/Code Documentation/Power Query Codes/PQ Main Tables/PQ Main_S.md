### Power Query code for Main_S table:
All data transformation steps are done in **Main_S**. 
The source table of **Main_S** is retrieved using **Path** (a dynamic project folder path) and **FileName** parameter.
To switch to another sector's **Main_S** table, choose a different **FileName** parameter value corresponding to the desired sector.
```m
let
    Source = Excel.CurrentWorkbook(){[Name="DynamicPath"]}[Content],
    Path = Source{0}[Path],
    Get_File = Csv.Document(File.Contents(Path & "Downloaded CSV Files\" & FileName),[Delimiter=",", Columns=10, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Removed Top Rows" = Table.Skip(Get_File,3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers",{"Last"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Removed Columns", {{"Market Cap", each Text.BeforeDelimiter(_, " "), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted Text Before Delimiter",{{"Market Cap", "Market Cap (M)"}, {"Free Cash Flow Per Share - Current (LTM)", "Free CF"}, {"Book Value Per Share - Current (LTM)", "BVPS"}, {"Earnings Per Share - TTM - Current (LTM)", "EPS"}, {"Return on Equity (ROE) - Current (LTM)", "ROE"}, {"Return on Assets (ROA) - Current (LTM)", "ROA"}, {"Financial Leverage (Assets/Equity) - Current (LTM)", "A/E"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","<empty>",null,Replacer.ReplaceValue,{"Free CF", "BVPS", "EPS", "ROE", "ROA", "A/E"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Replaced Value",{{"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"Market Cap (M)", Int64.Type}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"Free CF", each Number.Round(_, 2), type number}, {"BVPS", each Number.Round(_, 2), type number}, {"EPS", each Number.Round(_, 2), type number}, {"ROE", each Number.Round(_, 2), type number}, {"ROA", each Number.Round(_, 2), type number}, {"A/E", each Number.Round(_, 2), type number}}),
    #"Added Conditional Column" = Table.AddColumn(#"Rounded Off", "Market Cap Category", each if [#"Market Cap (M)"] <= 2000 then "Small" else if [#"Market Cap (M)"] < 10000 then "Mid" else if [#"Market Cap (M)"] >= 10000 then "Large" else null),
    #"Invoked Custom Function" = Table.AddColumn(#"Added Conditional Column", "Current Price", each #"Current Price"()),
    #"Invoked Custom Function1" = Table.AddColumn(#"Invoked Custom Function", "P/FCF", each #"P/FCF"()),
    #"Invoked Custom Function2" = Table.AddColumn(#"Invoked Custom Function1", "P/B", each #"P/B"()),
    #"Invoked Custom Function3" = Table.AddColumn(#"Invoked Custom Function2", "P/E", each #"P/E"()),
    #"Reordered Columns" = Table.ReorderColumns(#"Invoked Custom Function3",{"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "Market Cap Category"}),
    #"Grouped Rows" = Table.Group(#"Reordered Columns", {"Market Cap Category"}, {{"Table", each _, type table [Industry=nullable text, Symbol=nullable text, #"Market Cap (M)"=nullable number, Current Price=text, Free CF=number, BVPS=number, EPS=number, #"P/FCF"=text, #"P/B"=text, ROE=number, ROA=number, #"P/E"=text, #"A/E"=number, Market Cap Category=text]}}),
    Small = #"Grouped Rows"{[#"Market Cap Category"="Small"]}[Table],
    #"Removed Columns1" = Table.RemoveColumns(Small,{"Market Cap Category"})
in
    #"Removed Columns1"
```