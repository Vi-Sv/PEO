Sub SumMainCiphersOnly()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, j As Long
    Dim colLetter As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ws.Range(ws.Cells(2, 3), ws.Cells(2, lastCol)).NumberFormat = "General"
    
    For j = 3 To lastCol
        colLetter = Split(ws.Cells(1, j).Address, "$")(1)
        ws.Cells(2, j).Formula = "=SUMPRODUCT(ISNUMBER(A3:A" & lastRow & ")*" & colLetter & "3:" & colLetter & "3" & lastRow & ")"
        ws.Cells(2, j).Formula = Replace(ws.Cells(2, j).Formula, "3" & lastRow, CStr(lastRow))
    Next j
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
