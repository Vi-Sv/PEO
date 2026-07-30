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
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    With ws.Range(ws.Cells(2, 3), ws.Cells(2, lastCol))
        .NumberFormat = "General"
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    For j = 3 To lastCol
        colLetter = Split(ws.Cells(1, j).Address, "$")(1)
        ws.Cells(2, j).Formula = "=SUMPRODUCT(ISNUMBER($A$3:$A$" & lastRow & ")*" & colLetter & "$3:" & colLetter & "$" & lastRow & ")"
    Next j
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
