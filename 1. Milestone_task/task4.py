Sub ConvertUnitsToPieces_FinalSafe()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim unitStr As String
    Dim coef As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow
        ' Сброс коэффициента
        coef = 1
        
        ' Извлекаем Ед.изм, убираем пробелы, УНИЧТОЖАЕМ ТОЧКИ в памяти и переводим в нижний регистр
        unitStr = Trim(ws.Cells(i, 6).Value)
        unitStr = Replace(unitStr, " ", "")
        unitStr = Replace(unitStr, ".", "")
        unitStr = LCase(unitStr)
        
        ' Схлопываем варианты с точкой и без в единое правило
        If unitStr = "10шт" Then
            coef = 10
        ElseIf unitStr = "100шт" Then
            coef = 100
        ElseIf unitStr = "1000шт" Then
            coef = 1000
        End If
        
        ' Пересчет укрупненных позиций
        If coef > 1 Then
            If IsNumeric(ws.Cells(i, 3).Value) And ws.Cells(i, 3).Value <> "" Then
                ws.Cells(i, 3).Value = ws.Cells(i, 3).Value * coef
            End If
            
            If IsNumeric(ws.Cells(i, 4).Value) And ws.Cells(i, 4).Value <> "" Then
                ws.Cells(i, 4).Value = ws.Cells(i, 4).Value * coef
            End If
            
            If IsNumeric(ws.Cells(i, 5).Value) And ws.Cells(i, 5).Value <> "" Then
                ws.Cells(i, 5).Value = ws.Cells(i, 5).Value * coef
            End If
            
            ' Записываем в ячейку строгий эталон без точки
            ws.Cells(i, 6).Value = "шт"
            
            If IsNumeric(ws.Cells(i, 7).Value) And ws.Cells(i, 7).Value <> "" Then
                ws.Cells(i, 7).Value = ws.Cells(i, 7).Value / coef
            End If
            
        ' Защита: Если это была просто единица "шт." с точкой — убираем точку, сохраняя исходные числа
        ElseIf unitStr = "шт" Then
            ws.Cells(i, 6).Value = "шт"
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Контроль штук завершен! Все вариации приведены к единому стандарту 'шт'.", vbInformation, "Успех"
End Sub
