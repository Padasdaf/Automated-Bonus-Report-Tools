Sub StepFive()
    Dim otherWb As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim cell As Range
    Dim filePath As String
    
    filePath = "C:\Users\503144637\Documents\Copy of GEHZ direct mapping 1.xlsx"
    Set otherWb = Workbooks.Open(filePath)
    Set wsSource = otherWb.Sheets("Sheet2")
    
    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsDest = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsDest.Name = "Sheet4"
    
    For Each cell In wsDest.UsedRange
        If cell.HasFormula Then
            cell.Formula = Replace(cell.Formula, "[" & otherWb.Name & "]", "")
        End If
    Next cell
    
    otherWb.Close SaveChanges:=False
    
End Sub
