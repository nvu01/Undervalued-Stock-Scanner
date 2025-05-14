Attribute VB_Name = "Clear_Empty_Used_Range"
Sub ClearEmptyRows()
    Dim ws As Worksheet
    Dim dataLastRow As Long
    Dim usedLastRow As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array("Result_L", "Result_M", "Result_S")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Find the last row with actual data
        dataLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' Find the last row Excel considers part of the UsedRange
        usedLastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
        If dataLastRow < usedLastRow Then
            ws.Range(ws.Rows(dataLastRow + 1), ws.Rows(usedLastRow)).Delete
        End If
    Next sheetName
End Sub

Sub CheckUsedRangeExtraRows()
    Dim ws As Worksheet
    Dim dataLastRow As Long
    Dim usedLastRow As Long

    Set ws = ActiveSheet

    ' Find the last row with actual data
    dataLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find the last row Excel considers part of the UsedRange
    usedLastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row

    ' Compare
    If usedLastRow > dataLastRow Then
        MsgBox "UsedRange includes extra blank rows: " & (usedLastRow - dataLastRow), vbExclamation
    Else
        MsgBox "UsedRange matches data rows — no extra blank rows.", vbInformation
    End If
End Sub
