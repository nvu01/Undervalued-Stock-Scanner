## Extract Mean Tables
`Mean Tables` query is used to create `Extract Mean Tables` function. 
This function returns all the mean tables in a worksheet specified by a parameter.

Parameters: FolderPath, FileName

### Query: Mean Tables
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName),null,true),
    #"Removed Top Rows" = Table.Skip(Source,28),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Item", "Kind", "Hidden"})
in
    #"Removed Columns"
```

### Function: Extract Mean Tables
```
let
    Source = (FolderPath as any, FileName as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName),null,true),
    #"Removed Top Rows" = Table.Skip(Source,28),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Item", "Kind", "Hidden"})
in
    #"Removed Columns"
in
    Source
```

## Combine & Transform Mean Tables
`Extract Mean Tables` function was invoked in `Combined Mean Table` query to extract all mean tables in all worksheets in the `Undervalued Stock Scanner\Main\Results` folder. 

```
let
    Source = Folder.Files(FolderPath & "\Main\Results\"),
    #"Removed Columns" = Table.RemoveColumns(Source,{"Extension", "Date accessed", "Date created", "Attributes", "Folder Path"}),
    #"Added FolderPath" = Table.AddColumn(#"Removed Columns", "FolderPath", each FolderPath),
    #"Invoked Custom Function" = Table.AddColumn(#"Added FolderPath", "Extract Mean Tables", each #"Extract Mean Tables"([FolderPath], [Name])),
    #"Expanded Extract Mean Tables" = Table.ExpandTableColumn(#"Invoked Custom Function", "Extract Mean Tables", {"Name", "Data"}, {"Name.1", "Data"}),
    #"Expanded Data" = Table.ExpandTableColumn(#"Expanded Extract Mean Tables", "Data", {"Column1", "Column2", "Column3"}, {"Column1", "Column2", "Column3"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Data",{"Content", "FolderPath", "Column3"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns1",{{"Column2", type number}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type", {"Column2"}),
    #"Rounded Off" = Table.TransformColumns(#"Removed Errors",{{"Column2", each Number.Round(_, 2), type number}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Rounded Off", {{"Name", each Text.BeforeDelimiter(_, "."), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted Text Before Delimiter",{{"Name", "Sector"}, {"Name.1", "Metric Type"}, {"Column1", "Industry"}, {"Column2", "Mean"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Market Cap", each if Text.EndsWith([Metric Type], "L") then "Large" else if Text.EndsWith([Metric Type], "M") then "Mid" else "Small"),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Added Conditional Column", {{"Metric Type", each Text.BeforeDelimiter(_, "_", 1), type text}}),
    #"Pivoted Column" = Table.Pivot(#"Extracted Text Before Delimiter1", List.Distinct(#"Extracted Text Before Delimiter1"[#"Metric Type"]), "Metric Type", "Mean", List.Sum)
in
    #"Pivoted Column"
```