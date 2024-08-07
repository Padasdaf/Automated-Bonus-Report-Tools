Sub StepThree()
    Dim wsDest As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim lastRow As Long
    Dim wsPivot As Worksheet
    Dim columnHeaders As Variant
    Dim i As Long, j As Long
    Dim header As String
    
    Set wsDest = ThisWorkbook.Sheets("Sheet2")
    lastRow = wsDest.Cells(wsDest.Rows.Count, "C").End(xlUp).Row
    
    Set ptRange = wsDest.Range("C1:O" & lastRow)
    Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPivot.Name = "PivotTableSheet"
    
    Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsPivot.Range("A1"), TableName:="MyPivotTable")
    
    columnHeaders = Array("1", "2", "3", "4", "5", "6", "7", "8", "9")
    
    For j = 1 To wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
        If wsDest.Cells(1, j).Value = columnHeaders(0) Then
            pt.PivotFields(columnHeaders(0)).Orientation = xlRowField
            Exit For
        End If
    Next j
    
    For i = LBound(columnHeaders) To UBound(columnHeaders)
        If columnHeaders(i) <> columnHeaders(0) Then
            header = columnHeaders(i)
            For j = 1 To wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
                If wsDest.Cells(1, j).Value = header Then
                    pt.PivotFields(header).Orientation = xlDataField
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub
