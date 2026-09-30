Макрос идет по графе А, и для уровня 1 и Уровня 2 в следующих графах удаляет значения: BV BX CA CC CE CH CI CL CN CQ CS CU CX CY DB DD DG DI DK DN DO DR DT DW DY EA ED EE EH EJ EM EO EQ ET EU EX EZ FC FE FG FJ FK
-------
Sub ClearValuesForLevels1And2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim arrClear() As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Нарезка столбцов для очистки строго по вашему ТЗ
    arrClear = Split("BV,BX,CA,CC,CE,CH,CI,CL,CN,CQ,CS,CU,CX,CY,DB,DD,DG,DI,DK,DN,DO,DR,DT,DW,DY,EA,ED,EE,EH,EJ,EM,EO,EQ,ET,EU,EX,EZ,FC,FE,FG,FJ,FK", ",")
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' Проверяем, если строка является "Уровень 1" или "Уровень 2"
        If valA = "Уровень 1" Or valA = "Уровень 2" Then
            ' Очищаем содержимое ячеек в текущей строке для указанных столбцов
            For j = 0 To UBound(arrClear)
                ws.Cells(i, arrClear(j)).ClearContents
            Next j
        End If
    Next i
    
    ' Восстановление стандартных настроек Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
