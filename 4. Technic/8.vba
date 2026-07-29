Sub HideZerosWithFontColor()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim valA As String, valB As String, bgCol As Long
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 3 To lastRow
        valA = Trim(CStr(ws.Cells(i, "A").Value))
        valB = Trim(CStr(ws.Cells(i, "B").Value))
        
        If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
            bgCol = RGB(38, 38, 38)
        ElseIf valB <> "" And IsNumeric(valA) Then
            bgCol = RGB(43, 56, 75)
        Else
            bgCol = RGB(255, 255, 255)
        End If
        
        For j = 3 To lastCol
            With ws.Cells(i, j).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                .Font.Color = bgCol
            End With
        Next j
    Next i
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
