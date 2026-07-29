Sub DarkenSubgroupsAndSeparatorsFinal()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim valA As String, valB As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 3 To lastRow
        valA = Trim(CStr(ws.Cells(i, "A").Value))
        valB = Trim(CStr(ws.Cells(i, "B").Value))
        
        If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
            With ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol))
                .Interior.Color = RGB(38, 38, 38)
                .Font.Color = RGB(255, 255, 255)
            End With
        ElseIf valA = "" And valB = "" Then
            ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol)).Interior.Color = RGB(38, 38, 38)
        End If
    Next i
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
