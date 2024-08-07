Sub StepOne()
    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim destLastRow As Long
    Dim wsExists As Boolean
    
    Set ws = ThisWorkbook.Sheets("Recovered_Sheet1")
    
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Sheet2")
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "Sheet2"
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow
        If IsNumeric(ws.Cells(i, "A").Value) Then
            ws.Cells(i, "A").Value = ws.Cells(i, "A").Value * 1
        End If
    Next i
    
    For i = lastRow To 1 Step -1
        If ws.Cells(i, "E").Value = "indirect" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    destLastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
    ws.Range("A1:Q" & lastRow).Copy wsDest.Range("A" & destLastRow)
    
End Sub
