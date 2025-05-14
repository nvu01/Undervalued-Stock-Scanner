### Power Query Code for Transformation Builder:
All data transformation steps are done in **Transformation Builder**.  
**Transformation Builder** is used to create **Transformation Function**.  
The source table of **Transformation Builder** is retrieved using **Path** (a dynamic project folder path), **FileName** parameter, and **ResultSheet** parameter.
```m
let
    Source = Excel.CurrentWorkbook(){[Name="DynamicPath"]}[Content],
    Path = Source{0}[Path],
    Get_File = Excel.Workbook(File.Contents(Path & "Main\Results\" & FileName), null, true),
    Get_Sheet = Get_File{[Item=ResultSheet,Kind="Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Get_Sheet,5),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers",{"P/FCF Z Score", "P/B Z Score", "ROE Z Score", "ROA Z Score", "A/E Z Score", "P/E Z Score", "SCORE"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"Market Cap (M)", Int64.Type}, {"Symbol", type text}, {"Industry", type text}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"Free CF", each Number.Round(_, 2), type number}, {"BVPS", each Number.Round(_, 2), type number}, {"EPS", each Number.Round(_, 2), type number}, {"P/FCF", each Number.Round(_, 2), type number}, {"P/B", each Number.Round(_, 2), type number}, {"ROE", each Number.Round(_, 2), type number}, {"ROA", each Number.Round(_, 2), type number}, {"A/E", each Number.Round(_, 2), type number}, {"P/E", each Number.Round(_, 2), type number}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Rounded Off", {{"Industry", "None"}}),
    #"Removed Blank Rows" = Table.SelectRows(#"Replaced Errors", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))
in
    #"Removed Blank Rows"
```