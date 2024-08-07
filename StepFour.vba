Sub StepFour()
    Dim wsNew As Worksheet
    
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNew.Name = "NewSheet"
    
    wsNew.Range("A1").Value = 300003163
    wsNew.Range("A2").Value = 300003224
    wsNew.Range("A3").Value = 300003057
    wsNew.Range("A4").Value = 300003282
    wsNew.Range("A10").Value = 212810641
    wsNew.Range("A12").Value = 300003215
    
    wsNew.Range("B1").Value = "description1"
    wsNew.Range("B2").Value = "description2"
    wsNew.Range("B3").Value = "description3"
    wsNew.Range("B4").Value = "description4"
    wsNew.Range("B10").Value = "description5"
    wsNew.Range("B12").Value = "description6"
    
    wsNew.Range("C1").Value = "VBM"
    wsNew.Range("C2").Value = "VBM"
    wsNew.Range("C3").Value = "Fitter"
    
    wsNew.Range("D1").Value = "Du, Jianfeng"
    wsNew.Range("D2").Value = "Du, Jianfeng"
    wsNew.Range("D3").Value = "Feng, Hongjun"
    wsNew.Range("D4").Value = "Qian, Caihua"
    wsNew.Range("D10").Value = "Quality"
    wsNew.Range("D12").Value = "Quality"
    
    wsNew.Range("E1").Formula = "=VLOOKUP(A1,direct!A:G,7,0)"
    wsNew.Range("E2").Formula = "=VLOOKUP(A2,direct!A:G,7,0)"
    wsNew.Range("E3").Formula = "=VLOOKUP(A3,direct!A:G,7,0)"
    wsNew.Range("E4").Formula = "=VLOOKUP(A4,direct!A:G,7,0)"
    wsNew.Range("E10").Formula = "=VLOOKUP(A10,direct!A:G,7,0)"
    wsNew.Range("E12").Formula = "=VLOOKUP(A12,direct!A:G,7,0)"
    
End Sub
