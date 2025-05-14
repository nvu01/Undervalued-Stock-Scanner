## Transform Mean Table
`Sample_Mean_Table` query is used to create `Transform Mean Table` function. 
The function applies transformation steps to a specified mean table within a worksheet.

Parameters: FolderPath, FileName, MeanTable

### Query: Sample_Mean_Table
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Get Mean Table" = Source{[Item=MeanTable,Kind="DefinedName"]}[Data],
    #"Removed Top Rows" = Table.Skip(#"Get Mean Table",1),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Column3"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Column2", type number}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"Column2", each Number.Round(_, 2), type number}})
in
    #"Rounded Off"
```

### Function: Transform Mean Table
```
let
    Source = (FolderPath as any, FileName as any, MeanTable as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Get Mean Table" = Source{[Item=MeanTable,Kind="DefinedName"]}[Data],
    #"Removed Top Rows" = Table.Skip(#"Get Mean Table",1),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Column3"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Column2", type number}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"Column2", each Number.Round(_, 2), type number}})
in
    #"Rounded Off"
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
    #"Removed Top Rows" = Table.Skip(Source,28),
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
    #"Removed Top Rows" = Table.Skip(Source,28),
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

## Query: Mean_All
`Combine Mean Tables` function was invoked in this query for all the worksheets in `Undervalued Stock Scanner\Main\Results` folder. 
All the mean tables in all the worksheets are combined in one table and grouped by worksheet name and industry.
```
let
    Source = Folder.Files(FolderPath & "\Main\Results\"),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Date modified", "Folder Path"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Removed Other Columns", {{"Folder Path", each Text.BeforeDelimiter(_, "\M"), type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Extracted Text Before Delimiter", "Combine Mean Tables", each #"Combine Mean Tables"([Folder Path], [Name])),
    #"Expanded Combine Mean Tables" = Table.ExpandTableColumn(#"Invoked Custom Function", "Combine Mean Tables", {"Industry", "Mean_AE_L", "Mean_AE_M", "Mean_AE_S", "Mean_PB_L", "Mean_PB_M", "Mean_PB_S", "Mean_PE_L", "Mean_PE_M", "Mean_PE_S", "Mean_PFCF_L", "Mean_PFCF_M", "Mean_PFCF_S", "Mean_ROA_L", "Mean_ROA_M", "Mean_ROA_S", "Mean_ROE_L", "Mean_ROE_M", "Mean_ROE_S"}, {"Industry", "Mean_AE_L", "Mean_AE_M", "Mean_AE_S", "Mean_PB_L", "Mean_PB_M", "Mean_PB_S", "Mean_PE_L", "Mean_PE_M", "Mean_PE_S", "Mean_PFCF_L", "Mean_PFCF_M", "Mean_PFCF_S", "Mean_ROA_L", "Mean_ROA_M", "Mean_ROA_S", "Mean_ROE_L", "Mean_ROE_M", "Mean_ROE_S"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Combine Mean Tables",{"Date modified", "Folder Path"})
in
    #"Removed Columns"
```

## Transform Result Table
`Sample_Result_Table` query is used to create `Transform Result Table` function. 
The function applies transformation steps to a specified result table within a worksheet.

Parameters: FolderPath, FileName, MeanTable

### Query: Sample_Result_Table
```
let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    Result_Sheet = Source{[Item=ResultSheet,Kind="Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Result_Sheet,6),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Industry", type text}, {"Symbol", type text}, {"Market Cap (M)", Int64.Type}, {"Current Price", type number}, {"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"P/FCF Z-Score", type number}, {"P/B Z-Score", type number}, {"ROE Z-Score", type number}, {"ROA Z-Score", type number}, {"A/E Z-Score", type number}, {"P/E Z-Score", type number}, {"SCORE", Int64.Type}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"P/FCF", each Number.Round(_, 2), type number}, {"P/B", each Number.Round(_, 2), type number}, {"P/E", each Number.Round(_, 2), type number}, {"P/FCF Z-Score", each Number.Round(_, 2), type number}, {"P/B Z-Score", each Number.Round(_, 2), type number}, {"ROE Z-Score", each Number.Round(_, 2), type number}, {"ROA Z-Score", each Number.Round(_, 2), type number}, {"A/E Z-Score", each Number.Round(_, 2), type number}, {"P/E Z-Score", each Number.Round(_, 2), type number}}),
    #"Added Custom" = Table.AddColumn(#"Rounded Off", "Market Cap", each if Text.Contains(ResultSheet, "S") then "Small" else if Text.Contains(ResultSheet, "M") then "Mid" else "Large")
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
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"P/FCF", each Number.Round(_, 2), type number}, {"P/B", each Number.Round(_, 2), type number}, {"P/E", each Number.Round(_, 2), type number}, {"P/FCF Z-Score", each Number.Round(_, 2), type number}, {"P/B Z-Score", each Number.Round(_, 2), type number}, {"ROE Z-Score", each Number.Round(_, 2), type number}, {"ROA Z-Score", each Number.Round(_, 2), type number}, {"A/E Z-Score", each Number.Round(_, 2), type number}, {"P/E Z-Score", each Number.Round(_, 2), type number}}),
    #"Added Custom" = Table.AddColumn(#"Rounded Off", "Market Cap", each if Text.Contains(ResultSheet, "S") then "Small" else if Text.Contains(ResultSheet, "M") then "Mid" else "Large")
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
    #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",35),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Item", "Kind", "Hidden", "Data"}),
    #"Added Folder Path" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added File Name" = Table.AddColumn(#"Added Folder Path", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added File Name", "Transform Result Table", each #"Transform Result Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Result Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Result Table", {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}, {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Result Table",{"Name", "Folder Path"}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Removed Columns1", {"Industry"}),
    #"Added Criterion 1" = Table.AddColumn(#"Removed Errors", "Criterion 1", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_PB_L" else if [#"Market Cap"] = "Mid" then "Mean_PB_M" else "Mean_PB_S",
    average = Record.Field(row, mean_col)
in
    if [#"P/B"] < 0.7 * average then 1 else 0),

    #"Added Criterion 2" = Table.AddColumn(#"Added Criterion 1", "Criterion 2", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_ROE_L" else if [#"Market Cap"] = "Mid" then "Mean_ROE_M" else "Mean_ROE_S",
    average = Record.Field(row, mean_col)
in
    if [#"ROE"] > average then 1 else 0),

    #"Added Criterion 3" = Table.AddColumn(#"Added Criterion 2", "Criterion 3", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_ROA_L" else if [#"Market Cap"] = "Mid" then "Mean_ROA_M" else "Mean_ROA_S",
    average = Record.Field(row, mean_col)
in
    if [#"ROA"] > average then 1 else 0),

    #"Added Criterion 4" = Table.AddColumn(#"Added Criterion 3", "Criterion 4", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_PE_L" else if [#"Market Cap"] = "Mid" then "Mean_PE_M" else "Mean_PE_S",
    average = Record.Field(row, mean_col)
in
    if [#"P/E"] < average then 1 else 0),

    #"Added Criterion 5" = Table.AddColumn(#"Added Criterion 4", "Criterion 5", each if [#"P/E"] > 1 and [#"P/E"] < 25 then 1 else 0)
in
    #"Added Criterion 5"
```

### Function: Combine Result Tables
```
let
    Source = (FolderPath as any, FileName as any) => let
    Source = Excel.Workbook(File.Contents(FolderPath & "\Main\Results\" & FileName), null, true),
    #"Removed Top Rows" = Table.Skip(Source,8),
    #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",35),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Item", "Kind", "Hidden", "Data"}),
    #"Added Folder Path" = Table.AddColumn(#"Removed Columns", "Folder Path", each FolderPath),
    #"Added File Name" = Table.AddColumn(#"Added Folder Path", "File Name", each FileName),
    #"Invoked Custom Function" = Table.AddColumn(#"Added File Name", "Transform Result Table", each #"Transform Result Table"([Folder Path], [File Name], [Name])),
    #"Expanded Transform Result Table" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transform Result Table", {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}, {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Transform Result Table",{"Name", "Folder Path"}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Removed Columns1", {"Industry"}),
    #"Added Criterion 1" = Table.AddColumn(#"Removed Errors", "Criterion 1", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_PB_L" else if [#"Market Cap"] = "Mid" then "Mean_PB_M" else "Mean_PB_S",
    average = Record.Field(row, mean_col)
in
    if [#"P/B"] < 0.7 * average then 1 else 0),

    #"Added Criterion 2" = Table.AddColumn(#"Added Criterion 1", "Criterion 2", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_ROE_L" else if [#"Market Cap"] = "Mid" then "Mean_ROE_M" else "Mean_ROE_S",
    average = Record.Field(row, mean_col)
in
    if [#"ROE"] > average then 1 else 0),

    #"Added Criterion 3" = Table.AddColumn(#"Added Criterion 2", "Criterion 3", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_ROA_L" else if [#"Market Cap"] = "Mid" then "Mean_ROA_M" else "Mean_ROA_S",
    average = Record.Field(row, mean_col)
in
    if [#"ROA"] > average then 1 else 0),

    #"Added Criterion 4" = Table.AddColumn(#"Added Criterion 3", "Criterion 4", each let
    industry = [Industry],
    sector = [#"File Name"],
    row = Table.SelectRows(All_Means, each [Name] = sector and [Industry] = industry){0},
    mean_col = if [#"Market Cap"] = "Large" then "Mean_PE_L" else if [#"Market Cap"] = "Mid" then "Mean_PE_M" else "Mean_PE_S",
    average = Record.Field(row, mean_col)
in
    if [#"P/E"] < average then 1 else 0),

    #"Added Criterion 5" = Table.AddColumn(#"Added Criterion 4", "Criterion 5", each if [#"P/E"] > 1 and [#"P/E"] < 25 then 1 else 0)
in
    #"Added Criterion 5"
in
    Source
```

## Query: All_Results
`Combine Result Tables` function was invoked in this query for all the worksheets in `Undervalued Stock Scanner\Main\Results\` folder. 
All the result tables in all the worksheets are combined in one table and grouped by sector and industry.
```
let
    Source = Folder.Files(FolderPath & "\Main\Results\"),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Date modified", "Folder Path"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Removed Other Columns", {{"Folder Path", each Text.BeforeDelimiter(_, "\M"), type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Extracted Text Before Delimiter", "Combine Result Tables", each #"Combine Result Tables"([Folder Path], [Name])),
    #"Expanded Combine Result Tables" = Table.ExpandTableColumn(#"Invoked Custom Function", "Combine Result Tables", {"File Name", "Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap", "Criterion 1", "Criterion 2", "Criterion 3", "Criterion 4", "Criterion 5"}, {"File Name", "Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E", "P/FCF Z-Score", "P/B Z-Score", "ROE Z-Score", "ROA Z-Score", "A/E Z-Score", "P/E Z-Score", "SCORE", "Market Cap", "Criterion 1", "Criterion 2", "Criterion 3", "Criterion 4", "Criterion 5"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Combine Result Tables",{"File Name", "Folder Path"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Market Cap (M)", Int64.Type}, {"Current Price", type number}, {"Free CF", type number}, {"BVPS", type number}, {"EPS", type number}, {"P/FCF", type number}, {"P/B", type number}, {"ROE", type number}, {"ROA", type number}, {"A/E", type number}, {"P/E", type number}, {"P/FCF Z-Score", type number}, {"P/B Z-Score", type number}, {"ROE Z-Score", type number}, {"ROA Z-Score", type number}, {"A/E Z-Score", type number}, {"P/E Z-Score", type number}, {"Criterion 1", Int64.Type}, {"Criterion 2", Int64.Type}, {"Criterion 3", Int64.Type}, {"Criterion 4", Int64.Type}, {"Criterion 5", Int64.Type}, {"SCORE", Int64.Type}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Changed Type", {{"Name", each Text.BeforeDelimiter(_, "."), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted Text Before Delimiter1",{{"Name", "Sector"}})
in
    #"Renamed Columns"
```