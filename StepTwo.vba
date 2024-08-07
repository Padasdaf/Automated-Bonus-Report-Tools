Sub StepTwo()
    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim wsOther As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim otherWb As Workbook
    Dim matchRow As Long
    Dim filePath As String
    
    Set wsDest = ThisWorkbook.Sheets("direct")
    filePath = "C:\Users\503144637\Documents\Copy of GEHZ Direct mapping 1.xlsx"
    
    Set otherWb = Workbooks.Open(filePath)
    Set wsOther = otherWb.Sheets("Sheet1")
    
    wsDest.Columns("C").NumberFormat = "General"
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If IsNumeric(wsDest.Cells(i, "A").Value) Then
            matchRow = 0
            On Error Resume Next
            matchRow = Application.WorksheetFunction.Match(wsDest.Cells(i, "A").Value, wsOther.Columns("A"), 0)
            On Error GoTo 0
            
            If matchRow > 0 Then
                wsDest.Cells(i, "C").Value = wsOther.Cells(matchRow, "B").Value
            Else
                wsDest.Cells(i, "C").Value = "No Match Found"
            End If
        End If
    Next i
    
    otherWb.Close SaveChanges:=False
    
End Sub
