Sub StepThree()
    Dim wsDest As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim lastRow As Long
    Dim wsPivot As Worksheet
    Dim columnHeaders As Variant
    Dim i As Long, j As Long
    
    Set wsDest = ThisWorkbook.Sheets("Sheet2")
    lastRow = wsDest.Cells(wsDest.Rows.Count, "C").End(xlUp).Row
  
    Set ptRange = wsDest.Range("C1:O" & lastRow)
    Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPivot.Name = "PivotTableSheet"
    
    Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsPivot.Range("A1"), TableName:="MyPivotTable")
    
    columnHeaders = Array("Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7", "Header8", "Header9") ' Replace with actual headers
    
    For i = LBound(columnHeaders) To UBound(columnHeaders)
        For j = 1 To wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
            If wsDest.Cells(1, j).Value = columnHeaders(i) Then
                pt.PivotFields(columnHeaders(i)).Orientation = xlDataField
                Exit For
            End If
        Next j
    Next i
End Sub
