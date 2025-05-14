### Power Query Code for Quartile_S
All data transformation steps are done in **Quartile_S**. 
The source table of **Quartile_S** is retrieved using **Path** (a dynamic project folder path) and **FileName** parameter.
To switch to another sector's **Quartile_S** table, choose a different **FileName** parameter value corresponding to the desired sector.
```m
let
    Source = Excel.CurrentWorkbook(){[Name="DynamicPath"]}[Content],
    Path = Source{0}[Path],
    Get_File = Csv.Document(File.Contents(Path & "Downloaded CSV Files\" & FileName),[Delimiter=",", Columns=10, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Removed Top Rows" = Table.Skip(Get_File,3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Grouped Rows" = Table.Group(#"Promoted Headers", {"Industry"}, {{"Table", each _, type table [Industry=nullable text, Symbol=nullable text, Market Cap=nullable text, Last=nullable text, #"Free Cash Flow Per Share - Current (LTM)"=nullable text, #"Book Value Per Share - Current (LTM)"=nullable text, #"Earnings Per Share - TTM - Current (LTM)"=nullable text, #"Return on Equity (ROE) - Current (LTM)"=nullable text, #"Return on Assets (ROA) - Current (LTM)"=nullable text, #"Financial Leverage (Assets/Equity) - Current (LTM)"=nullable text]}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Table"}),
    #"Invoked Custom Function" = Table.AddColumn(#"Removed Columns", "P/FCF LB", each #"P/FCF LB"()),
    #"Invoked Custom Function1" = Table.AddColumn(#"Invoked Custom Function", "P/FCF UB", each #"P/FCF UB"()),
    #"Invoked Custom Function2" = Table.AddColumn(#"Invoked Custom Function1", "P/B LB", each #"P/B LB"()),
    #"Invoked Custom Function3" = Table.AddColumn(#"Invoked Custom Function2", "P/B UB", each #"P/B UB"()),
    #"Invoked Custom Function4" = Table.AddColumn(#"Invoked Custom Function3", "ROE LB", each #"ROE LB"()),
    #"Invoked Custom Function5" = Table.AddColumn(#"Invoked Custom Function4", "ROE UB", each #"ROE UB"()),
    #"Invoked Custom Function6" = Table.AddColumn(#"Invoked Custom Function5", "ROA LB", each #"ROA LB"()),
    #"Invoked Custom Function7" = Table.AddColumn(#"Invoked Custom Function6", "ROA UB", each #"ROA UB"()),
    #"Invoked Custom Function8" = Table.AddColumn(#"Invoked Custom Function7", "P/E LB", each #"P/E LB"()),
    #"Invoked Custom Function9" = Table.AddColumn(#"Invoked Custom Function8", "P/E UB", each #"P/E UB"()),
    #"Invoked Custom Function10" = Table.AddColumn(#"Invoked Custom Function9", "A/E LB", each #"A/E LB"()),
    #"Invoked Custom Function11" = Table.AddColumn(#"Invoked Custom Function10", "A/E UB", each #"A/E UB"()),
    #"Invoked Custom Function12" = Table.AddColumn(#"Invoked Custom Function11", "P/FCF Q1", each #"S_P/FCF Q1"()),
    #"Invoked Custom Function13" = Table.AddColumn(#"Invoked Custom Function12", "P/FCF Q3", each #"S_P/FCF Q3"()),
    #"Invoked Custom Function14" = Table.AddColumn(#"Invoked Custom Function13", "P/B Q1", each #"S_P/B Q1"()),
    #"Invoked Custom Function15" = Table.AddColumn(#"Invoked Custom Function14", "P/B Q3", each #"S_P/B Q3"()),
    #"Invoked Custom Function16" = Table.AddColumn(#"Invoked Custom Function15", "ROE Q1", each #"S_ROE Q1"()),
    #"Invoked Custom Function17" = Table.AddColumn(#"Invoked Custom Function16", "ROE Q3", each #"S_ROE Q3"()),
    #"Invoked Custom Function18" = Table.AddColumn(#"Invoked Custom Function17", "ROA Q1", each #"S_ROA Q1"()),
    #"Invoked Custom Function19" = Table.AddColumn(#"Invoked Custom Function18", "ROA Q3", each #"S_ROA Q3"()),
    #"Invoked Custom Function20" = Table.AddColumn(#"Invoked Custom Function19", "P/E Q1", each #"S_P/E Q1"()),
    #"Invoked Custom Function21" = Table.AddColumn(#"Invoked Custom Function20", "P/E Q3", each #"S_P/E Q3"()),
    #"Invoked Custom Function22" = Table.AddColumn(#"Invoked Custom Function21", "A/E Q1", each #"S_A/E Q1"()),
    #"Invoked Custom Function23" = Table.AddColumn(#"Invoked Custom Function22", "A/E Q3", each #"S_A/E Q3"()),
    #"Reordered Columns" = Table.ReorderColumns(#"Invoked Custom Function23",{"Industry", "P/FCF Q1", "P/FCF Q3", "P/FCF LB", "P/FCF UB", "P/B Q1", "P/B Q3", "P/B LB", "P/B UB", "ROE Q1", "ROE Q3", "ROE LB", "ROE UB", "ROA Q1", "ROA Q3", "ROA LB", "ROA UB", "P/E Q1", "P/E Q3", "P/E LB", "P/E UB", "A/E Q1", "A/E Q3", "A/E LB", "A/E UB"})
in
    #"Reordered Columns"
```