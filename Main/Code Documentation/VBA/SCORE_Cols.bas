Attribute VB_Name = "SCORE_Cols"
Sub SCORE_Result_L()
    Dim ws As Worksheet
    Dim formulaCell As Range
    Dim fillRange As Range
    Dim colLastRow As Long
    Dim dataLastRow As Long
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Result_L")

    ' Starting formula cell
    Set formulaCell = ws.Range("T9")

    ' Determine the last used row in that column
    colLastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row

    ' Determine the last used row of actual data
    dataLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set the fill range from row 9 down to lastRow
    Set fillRange = ws.Range("T9:T" & dataLastRow)

    ' Autofill from T9 to last data row
    If dataLastRow > 9 Then
        ws.Range("T9").AutoFill Destination:=ws.Range("T9:T" & dataLastRow)
    End If
    
    ' Clear any extra values below the last row with data
    If colLastRow > dataLastRow Then
        ws.Range("T" & dataLastRow + 1 & ":T" & colLastRow).ClearContents
    End If
End Sub

Sub SCORE_Result_M()
    Dim ws As Worksheet
    Dim formulaCell As Range
    Dim fillRange As Range
    Dim colLastRow As Long
    Dim dataLastRow As Long
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Result_M")

    ' Starting formula cell
    Set formulaCell = ws.Range("T9")

    ' Determine the last used row in that column
    colLastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row

    ' Determine the last used row of actual data
    dataLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set the fill range from row 9 down to lastRow
    Set fillRange = ws.Range("T9:T" & dataLastRow)

    ' Autofill from T9 to last data row
    If dataLastRow > 9 Then
        ws.Range("T9").AutoFill Destination:=ws.Range("T9:T" & dataLastRow)
    End If
    
    ' Clear any extra values below the last row with data
    If colLastRow > dataLastRow Then
        ws.Range("T" & dataLastRow + 1 & ":T" & colLastRow).ClearContents
    End If
End Sub

Sub SCORE_Result_S()
    Dim ws As Worksheet
    Dim formulaCell As Range
    Dim fillRange As Range
    Dim colLastRow As Long
    Dim dataLastRow As Long
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Result_S")

    ' Starting formula cell
    Set formulaCell = ws.Range("T9")

    ' Determine the last used row in that column
    colLastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row

    ' Determine the last used row of actual data
    dataLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set the fill range from row 9 down to lastRow
    Set fillRange = ws.Range("T9:T" & dataLastRow)

    ' Autofill from T9 to last data row
    If dataLastRow > 9 Then
        ws.Range("T9").AutoFill Destination:=ws.Range("T9:T" & dataLastRow)
    End If
    
    ' Clear any extra values below the last row with data
    If colLastRow > dataLastRow Then
        ws.Range("T" & dataLastRow + 1 & ":T" & colLastRow).ClearContents
    End If
End Sub
