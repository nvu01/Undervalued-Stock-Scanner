### Power Query Code for All_Results table:
**All_Results** table is created by appending **Result_L_AllSectors**, **Result_M_AllSectors**, **Result_S_AllSectors**.
```m
let
    Source = Table.Combine({Result_L_AllSectors, Result_M_AllSectors, Result_S_AllSectors}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(Source, {{"Name", each Text.BeforeDelimiter(_, "."), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted Text Before Delimiter",{{"Name", "Sector"}})
in
    #"Renamed Columns"
```