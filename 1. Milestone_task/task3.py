Sub ConvertUnitsToM3_StrictSafe()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim unitStr As String
    Dim coef As Double
    
    ' Работаем строго с текущим активным листом Excel
    Set ws = ActiveSheet
    
    ' Находим последнюю заполненную строку по Столбцу 1 (А)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Выключаем обновление экрана для моментального расчета
    Application.ScreenUpdating = False
    
    ' Начинаем со 2-й строки, так как 1-я — это шапка таблицы
    For i = 2 To lastRow
        ' Важнейший сброс: для каждой новой строки коэффициент равен 1
        coef = 1
        
        ' Берем текст из Столбца 6 (Ед.изм.), убираем пробелы и переводим в строчные буквы
        unitStr = Trim(ws.Cells(i, 6).Value)
        unitStr = Replace(unitStr, " ", "")
        unitStr = LCase(unitStr)
        
        ' Жесткая проверка на абсолютное совпадение текстовых строк под М3
        If unitStr = "10м3" Then
            coef = 10
        ElseIf unitStr = "100м3" Then
            coef = 100
        ElseIf unitStr = "1000м3" Then
            coef = 1000
        End If
        
        ' Точечный пересчет запускается ТОЛЬКО если найден укрупненный коэффициент
        If coef > 1 Then
            ' Столбец 3: Объем (Исходный) -> Умножаем на коэффициент
            If IsNumeric(ws.Cells(i, 3).Value) And ws.Cells(i, 3).Value <> "" Then
                ws.Cells(i, 3).Value = ws.Cells(i, 3).Value * coef
            End If
            
            ' Столбец 4: Вып. до 04.2026 -> Умножаем на коэффициент
            If IsNumeric(ws.Cells(i, 4).Value) And ws.Cells(i, 4).Value <> "" Then
                ws.Cells(i, 4).Value = ws.Cells(i, 4).Value * coef
            End If
            
            ' Столбец 5: План на объем Begin -> Умножаем на коэффициент
            If IsNumeric(ws.Cells(i, 5).Value) And ws.Cells(i, 5).Value <> "" Then
                ws.Cells(i, 5).Value = ws.Cells(i, 5).Value * coef
            End If
            
            ' Столбец 6: Переименовываем укрупненную единицу в чистые м3
            ws.Cells(i, 6).Value = "м3"
            
            ' Столбец 7: Норма расхода ч/ч -> Жестко делим на коэффициент
            If IsNumeric(ws.Cells(i, 7).Value) And ws.Cells(i, 7).Value <> "" Then
                ws.Cells(i, 7).Value = ws.Cells(i, 7).Value / coef
            End If
        End If
    Next i
    
    ' Возвращаем обновление экрана в исходное состояние
    Application.ScreenUpdating = True
    
    MsgBox "Контроль кубических метров завершен! Строки с обычными м3 полностью защищены от изменений.", vbInformation, "Успех"
End Sub
