Attribute VB_Name = "FullRefresh"
Sub DataRefresh()
    ThisWorkbook.RefreshAll
    MsgBox "Refreshing data... please wait at least 30 seconds.", vbInformation
    Application.OnTime Now + TimeValue("00:00:15"), "ConvertToFormulas"
    Application.OnTime Now + TimeValue("00:00:40"), "RefreshPivotTables"
    Application.OnTime Now + TimeValue("00:00:41"), "SCORE_Result_L"
    Application.OnTime Now + TimeValue("00:00:41"), "SCORE_Result_M"
    Application.OnTime Now + TimeValue("00:00:41"), "SCORE_Result_S"
    Application.OnTime Now + TimeValue("00:00:42"), "ClearEmptyRows"
    Application.OnTime Now + TimeValue("00:00:42"), "ShowMessage"
End Sub

Sub ConvertToFormulas()
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim cell As Range
    Dim tableName As Variant
    Dim colName As Variant
    Dim colRange As Range

    ' Define tables and columns to process
    Dim tableConfig As Object
    Set tableConfig = CreateObject("Scripting.Dictionary")
    ' Format: tableName = Array(SheetName, Array(ColumnNames))
    tableConfig.Add "Main_L", Array("Main_L", Array("Current Price", "P/FCF", "P/B", "P/E"))
    tableConfig.Add "Main_M", Array("Main_M", Array("Current Price", "P/FCF", "P/B", "P/E"))
    tableConfig.Add "Main_S", Array("Main_S", Array("Current Price", "P/FCF", "P/B", "P/E"))
    tableConfig.Add "Quartile_L", Array("Quartile Tables", Array("P/FCF Q1", "P/FCF Q3", "P/FCF LB", "P/FCF UB", "P/B Q1", "P/B Q3", "P/B LB", "P/B UB", "ROE Q1", "ROE Q3", "ROE LB", "ROE UB", "ROA Q1", "ROA Q3", "ROA LB", "ROA UB", "A/E Q1", "A/E Q3", "A/E LB", "A/E UB", "P/E Q1", "P/E Q3", "P/E LB", "P/E UB"))
    tableConfig.Add "Quartile_M", Array("Quartile Tables", Array("P/FCF Q1", "P/FCF Q3", "P/FCF LB", "P/FCF UB", "P/B Q1", "P/B Q3", "P/B LB", "P/B UB", "ROE Q1", "ROE Q3", "ROE LB", "ROE UB", "ROA Q1", "ROA Q3", "ROA LB", "ROA UB", "A/E Q1", "A/E Q3", "A/E LB", "A/E UB", "P/E Q1", "P/E Q3", "P/E LB", "P/E UB"))
    tableConfig.Add "Quartile_S", Array("Quartile Tables", Array("P/FCF Q1", "P/FCF Q3", "P/FCF LB", "P/FCF UB", "P/B Q1", "P/B Q3", "P/B LB", "P/B UB", "ROE Q1", "ROE Q3", "ROE LB", "ROE UB", "ROA Q1", "ROA Q3", "ROA LB", "ROA UB", "A/E Q1", "A/E Q3", "A/E LB", "A/E UB", "P/E Q1", "P/E Q3", "P/E LB", "P/E UB"))

    ' Loop through each defined table and column
    For Each tableName In tableConfig.Keys
        Dim sheetName As String
        sheetName = tableConfig(tableName)(0)
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        On Error Resume Next

        Set tbl = ws.ListObjects(tableName)

        For Each colName In tableConfig(tableName)(1)
            On Error Resume Next
            Set colRange = tbl.ListColumns(colName).DataBodyRange
            For Each cell In colRange
                If Not cell.HasFormula Then
                    If Left(cell.Value, 1) = "=" Then
                        cell.Formula2 = cell.Value
                    End If
                End If
            Next cell
        Next colName
    Next tableName
End Sub

Sub RefreshPivotTables()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array("Mean_L", "Mean_M", "Mean_S")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next sheetName
End Sub

Sub ShowMessage()
    MsgBox "Data has been refreshed.", vbInformation
End Sub
