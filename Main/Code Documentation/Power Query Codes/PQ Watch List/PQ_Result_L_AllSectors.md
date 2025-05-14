### Power Query Code for Results_L_AllSectors table:
The dynamic path for **Undervalued Stock Scanner** folder is in the **DynamicPath** table, within the **Dynamic Path** sheet.  
**Transformation Function** is used to apply data transformation to a list of all Excel files in the **Results** folder.  
**FileName** parameter specifies the Excel files. **ResultSheet** parameter specifies which result sheet to apply data transformation. In this code, **Result_L** is used as argument for **ResultSheet**.
```m
let
    Source = Excel.CurrentWorkbook(){[Name="DynamicPath"]}[Content],
    Path = Source{0}[Path],
    Files = Folder.Files(Path & "Main\Results"),
    #"Removed Other Columns" = Table.SelectColumns(Files,{"Name"}),
    #"Invoked Custom Function" = Table.AddColumn(#"Removed Other Columns", "Transformation", each Transformation([Name], "Result_L")),
    #"Expanded Transformation" = Table.ExpandTableColumn(#"Invoked Custom Function", "Transformation", {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E"}, {"Industry", "Symbol", "Market Cap (M)", "Current Price", "Free CF", "BVPS", "EPS", "P/FCF", "P/B", "ROE", "ROA", "A/E", "P/E"})
in
    #"Expanded Transformation"
```