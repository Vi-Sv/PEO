Sub FormatReportSheetFinalWithBorders()
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
    
    ' Форматирование шапки (строка 1)
    With ws.Range(ws.Cells(1, "A"), ws.Cells(1, lastCol))
        .Interior.Color = RGB(43, 56, 75)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    For i = 1 To lastRow
        valA = Trim(CStr(ws.Cells(i, "A").Value))
        valB = Trim(CStr(ws.Cells(i, "B").Value))
        
        If i >= 3 Then
            If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
                With ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol))
                    .Interior.Color = RGB(255, 255, 255)
                    .Font.Color = RGB(0, 0, 0)
                    .Font.Bold = True
                    .Font.Italic = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            ElseIf valB <> "" And IsNumeric(valA) Then
                ws.Rows(i).RowHeight = 25
                With ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol))
                    .Interior.Color = RGB(43, 56, 75)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                    .Font.Italic = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            ElseIf valA = "" And valB = "" Then
                ws.Rows(i).RowHeight = 8
                ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol)).Interior.Color = RGB(200, 200, 200)
            End If
        End If
        
        ' Настройка границ для всех заполненных ячеек строки, исключая тонкие пустые разделители
        If i < 3 Or (valA <> "" Or valB <> "") Then
            With ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(128, 128, 128)
            End With
        Else
            ws.Range(ws.Cells(i, "A"), ws.Cells(i, lastCol)).Borders.LineStyle = xlNone
        End If
    Next i
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
