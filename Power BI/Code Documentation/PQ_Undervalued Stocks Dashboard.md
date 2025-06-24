# Mean Tables
## Transform Mean Table
`Sample_Mean_Table` query is used to create `Transform Mean Table` function. 
The function applies transformation steps to a specified mean table within a worksheet.

Parameters: FolderPath, FileName, MeanTable

### Query: Sample_Mean_Table
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Get Mean Table" = Source{[Item=MeanTable,Kind="DefinedName"]}[Data],
    #"Removed Columns" = Table.RemoveColumns(#"Get Mean Table",{"Column3"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Columns",1)
in
    #"Removed Top Rows"
```

### Function: Transform Mean Table
```
let
    Source = (FolderPath as any, FileName as any, MeanTable as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Get Mean Table" = Source{[Item=MeanTable,Kind="DefinedName"]}[Data],
    #"Removed Columns" = Table.RemoveColumns(#"Get Mean Table",{"Column3"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Columns",1)
in
    #"Removed Top Rows"
in
    Source
```

## Combine Mean Tables
`Industry_Means` query is used to create `Combine Mean Tables` function. 
The function applies `Transform Mean Table` function to all mean tables within a worksheet and combines them to create a table grouped by Industry.

Parameters: FolderPath, FileName

### Query: Industry_Means
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Removed Top Rows" = Table.Skip(Source,27),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Data", "Item", "Kind", "Hidden"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added Custom1", "Transform Mean Table", each #"Transform Mean Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Mean Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Mean Table", {"Column1", "Column2"}, {"Column1", "Column2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Mean Table",{"Folder Path", "File Name"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns1", List.Distinct(#"Removed Columns1"[Name]), "Name", "Column2", List.Sum),
    #"Renamed Columns" = Table.RenameColumns(#"Pivoted Column",{{"Column1", "Industry"}})
in
    #"Renamed Columns"
```

### Function: Combine Mean Tables
```
let
    Source = (FolderPath as any, FileName as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Removed Top Rows" = Table.Skip(Source,27),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Data", "Item", "Kind", "Hidden"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added Custom1", "Transform Mean Table", each #"Transform Mean Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Mean Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Mean Table", {"Column1", "Column2"}, {"Column1", "Column2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Mean Table",{"Folder Path", "File Name"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns1", List.Distinct(#"Removed Columns1"[Name]), "Name", "Column2", List.Sum),
    #"Renamed Columns" = Table.RenameColumns(#"Pivoted Column",{{"Column1", "Industry"}})
in
    #"Renamed Columns"
in
    Source
```

## Result Query: All_Means
`Combine Mean Tables` function was invoked in this query for all the worksheets in `Undervalued Stock Scanner\Main\Results` folder. 
All the mean tables in all the worksheets are combined in one table and grouped by worksheet name and industry.
```
let
    Source = Folder.Files(FolderPath & "\Main\Results\"),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Date modified", "Folder Path"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Removed Other Columns", {{"Folder Path", each Text.BeforeDelimiter(_, "\M"), type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Extracted Text Before Delimiter", "Combine Mean Tables", each #"Combine Mean Tables"([Folder Path], [Name])),
    #"Expanded Combine Mean Tables" = Table.ExpandTableColumn(#"Invoked Custom Function", "Combine Mean Tables", {"Industry", "Mean_AE_L", "Mean_AE_M", "Mean_AE_S", "Mean_PB_L", "Mean_PB_M", "Mean_PB_S", "Mean_PE_L", "Mean_PE_M", "Mean_PE_S", "Mean_PFCF_L", "Mean_PFCF_M", "Mean_PFCF_S", "Mean_ROA_L", "Mean_ROA_M", "Mean_ROA_S", "Mean_ROE_L", "Mean_ROE_M", "Mean_ROE_S"}, {"Industry", "Mean_AE_L", "Mean_AE_M", "Mean_AE_S", "Mean_PB_L", "Mean_PB_M", "Mean_PB_S", "Mean_PE_L", "Mean_PE_M", "Mean_PE_S", "Mean_PFCF_L", "Mean_PFCF_M", "Mean_PFCF_S", "Mean_ROA_L", "Mean_ROA_M", "Mean_ROA_S", "Mean_ROE_L", "Mean_ROE_M", "Mean_ROE_S"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Combine Mean Tables",{"Date modified", "Folder Path"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Name", "Sector"}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Renamed Columns", {{"Sector", each Text.BeforeDelimiter(_, "."), type text}})
in
    #"Extracted Text Before Delimiter1"
```

# Result Tables

## Transform Result Table
`Sample_Result_Table` query is used to create `Transform Result Table` function. 
The function applies transformation steps to a specified result table within a worksheet.

Parameters: FolderPath, FileName, MeanTable

### Query: Sample_Result_Table
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    Result_Sheet = Source{[Item=ResultSheet,Kind="Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Result_Sheet,7),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Industry", type text}, {"Symbol", type text}, {"Market Cap (M)", Int64.Type}, {"Current Price", type number}, {"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"P/FCF Z-Score", type number}, {"P/B Z-Score", type number}, {"ROE Z-Score", type number}, {"ROA Z-Score", type number}, {"A/E Z-Score", type number}, {"P/E Z-Score", type number}, {"SCORE", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Market Cap", each if Text.Contains(ResultSheet, "S") then "Small" else if Text.Contains(ResultSheet, "M") then "Mid" else "Large")
in
    #"Added Custom"
```

### Function: Transform Result Table
```
let
    Source = (FolderPath as any, FileName as any, ResultSheet as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    Result_Sheet = Source{[Item=ResultSheet,Kind="Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Result_Sheet,7),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Industry", type text}, {"Symbol", type text}, {"Market Cap (M)", Int64.Type}, {"Current Price", type number}, {"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"P/FCF Z-Score", type number}, {"P/B Z-Score", type number}, {"ROE Z-Score", type number}, {"ROA Z-Score", type number}, {"A/E Z-Score", type number}, {"P/E Z-Score", type number}, {"SCORE", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Market Cap", each if Text.Contains(ResultSheet, "S") then "Small" else if Text.Contains(ResultSheet, "M") then "Mid" else "Large")
in
    #"Added Custom"
in
    Source
```
## Combine Result Tables
`Industry_Results` query is used to create `Combine Result Tables` function. 
The function applies `Transform Result Table` function to all result tables within a worksheet and combines them to create a table grouped by Industry.
The function also creates 5 new Boolean columns based on some conditions that use data from `Mean_All` table.

Parameters: FolderPath, FileName

### Query: Industry_Results
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Removed Top Rows" = Table.Skip(Source,8),
    #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",34),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Item", "Kind", "Hidden", "Data"}),
    #"Added Folder Path" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added File Name" = Table.AddColumn(#"Added Folder Path", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added File Name", "Transform Result Table", each #"Transform Result Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Result Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Result Table", {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}, {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Result Table",{"Name", "Folder Path"}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Removed Columns1", {"Industry"})
in
    #"Removed Errors"
```

### Function: Combine Result Tables
```
let
    Source = (FolderPath as any, FileName as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Removed Top Rows" = Table.Skip(Source,8),
    #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",34),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Item", "Kind", "Hidden", "Data"}),
    #"Added Folder Path" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added File Name" = Table.AddColumn(#"Added Folder Path", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added File Name", "Transform Result Table", each #"Transform Result Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Result Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Result Table", {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}, {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Result Table",{"Name", "Folder Path"}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Removed Columns1", {"Industry"})
in
    #"Removed Errors"
in
    Source
```

## Result Query: All_Results
`Combine Result Tables` function was invoked in this query for all the worksheets in `Undervalued Stock Scanner\Main\Results\` folder. 
All the result tables in all the worksheets are combined in one table and grouped by sector and industry.
```
let
    Source = Folder.Files(FolderPath & "\Main\Results\"),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Date modified", "Folder Path"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Removed Other Columns", {{"Folder Path", each Text.BeforeDelimiter(_, "\M"), type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Extracted Text Before Delimiter", "Combine Result Tables", each #"Combine Result Tables"([Folder Path], [Name])),
    #"Expanded Combine Result Tables" = Table.ExpandTableColumn(#"Invoked Custom Function", "Combine Result Tables", {"File Name", "Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}, {"File Name", "Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Combine Result Tables",{"Folder Path", "File Name"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Market Cap (M)", Int64.Type}, {"Current Price", type number}, {"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"P/FCF Z-Score", type number}, {"P/B Z-Score", type number}, {"ROE Z-Score", type number}, {"ROA Z-Score", type number}, {"A/E Z-Score", type number}, {"P/E Z-Score", type number}, {"SCORE", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Name", "Sector"}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Renamed Columns", {{"Sector", each Text.BeforeDelimiter(_, "."), type text}})
in
    #"Extracted Text Before Delimiter1"
```

