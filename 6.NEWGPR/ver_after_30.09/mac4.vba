макрос идет по графе А, если Уровень 1, идет в графу I, запоминает формулу и передает в следующие графы: 
BY CF CO CV DE DL DU EB EK ER FA FH 
-------
Sub CopyLaborFormulaLevel1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim fI As String
    Dim targetCols() As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Массив целевых столбцов для переноса формулы
    targetCols = Split("BY,CF,CO,CV,DE,DL,DU,EB,EK,ER,FA,FH", ",")
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' Проверяем, если строка является "Уровень 1"
        If valA = "Уровень 1" Then
            ' Запоминаем локальную формулу из графы I текущей строки
            fI = ws.Cells(i, "I").FormulaR1C1
            
            ' Распространяем формулу по указанным столбцам
            For j = 0 To UBound(targetCols)
                ws.Cells(i, targetCols(j)).FormulaR1C1 = fI
            Next j
        End If
    Next i
    
    ' Восстановление стандартных настроек Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
