Sub InsertWorkSubgroupsOptimized()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = lastRow To 3 Step -1
        If Trim(ws.Cells(i, "B").Value) <> "" Then
            ws.Rows(i + 1 & ":" & i + 6).Insert Shift:=xlDown
            
            ws.Cells(i + 1, "A").Value = "М"
            ws.Cells(i + 1, "B").Value = "Монтажные работы"
            
            ws.Cells(i + 2, "A").Value = "Д"
            ws.Cells(i + 2, "B").Value = "Демонтажные работы"
            
            ws.Cells(i + 3, "A").Value = "П"
            ws.Cells(i + 3, "B").Value = "Погрузка, перевозка"
            
            ws.Cells(i + 4, "A").Value = "Л"
            ws.Cells(i + 4, "B").Value = "Перевозка людей"
            
            ws.Cells(i + 5, "A").Value = "З"
            ws.Cells(i + 5, "B").Value = "Земляные работы"
            
            ws.Rows(i + 6).RowHeight = 8
            
            ws.Rows(i + 1 & ":" & i + 5).Group
        End If
    Next i
    
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.ShowLevels RowLevels:=1
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
