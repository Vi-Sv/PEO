Sub CreateHierarchyTree()
    Dim srcWb As Workbook, destWb As Workbook
    Dim srcWs As Worksheet, destWs As Worksheet
    Dim dict As Object, lastRow As Long, i As Long
    Dim bVal As String, eVal As String, fVal As String, gVal As String
    Dim keyB As Variant, keyE As Variant, keyF As Variant, keyG As Variant
    Dim r As Long
    
    Set srcWb = ActiveWorkbook
    On Error Resume Next
    Set srcWs = srcWb.Sheets("ВР")
    On Error GoTo 0
    
    If srcWs Is Nothing Then
        MsgBox "Ошибка: Лист с именем 'ВР' не найден в активной книге!", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = srcWs.Cells(srcWs.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Ошибка: Нет данных для обработки в столбце B листа 'ВР'!", vbCritical
        Exit Sub
    End If
    
    For i = 2 To lastRow
        If Not IsError(srcWs.Cells(i, "B").Value) And _
           Not IsError(srcWs.Cells(i, "E").Value) And _
           Not IsError(srcWs.Cells(i, "F").Value) And _
           Not IsError(srcWs.Cells(i, "G").Value) Then
           
            bVal = Trim(CStr(srcWs.Cells(i, "B").Value))
            eVal = Trim(CStr(srcWs.Cells(i, "E").Value))
            fVal = Trim(CStr(srcWs.Cells(i, "F").Value))
            gVal = Trim(CStr(srcWs.Cells(i, "G").Value))
            
            If bVal <> "" Then
                If Not dict.Exists(bVal) Then Set dict(bVal) = CreateObject("Scripting.Dictionary")
                If eVal <> "" Then
                    If Not dict(bVal).Exists(eVal) Then Set dict(bVal)(eVal) = CreateObject("Scripting.Dictionary")
                    If fVal <> "" Then
                        If Not dict(bVal)(eVal).Exists(fVal) Then Set dict(bVal)(eVal)(fVal) = CreateObject("Scripting.Dictionary")
                        If gVal <> "" Then
                            dict(bVal)(eVal)(fVal)(gVal) = True
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    If dict.Count = 0 Then
        MsgBox "Ошибка: Не удалось собрать уникальные группы. Проверьте заполнение столбца B.", vbCritical
        Exit Sub
    End If
    
    Set destWb = Workbooks.Add
    Set destWs = destWb.Sheets(1)
    destWs.Outline.SummaryRow = xlSummaryAbove
    r = 1
    
    Application.ScreenUpdating = False
    For Each keyB In dict.Keys
        destWs.Cells(r, 1).Value = keyB
        destWs.Rows(r).OutlineLevel = 1
        r = r + 1
        For Each keyE In dict(keyB).Keys
            destWs.Cells(r, 1).Value = keyE
            destWs.Rows(r).OutlineLevel = 2
            r = r + 1
            For Each keyF In dict(keyB)(keyE).Keys
                destWs.Cells(r, 1).Value = keyF
                destWs.Rows(r).OutlineLevel = 3
                r = r + 1
                For Each keyG In dict(keyB)(keyE)(keyF).Keys
                    destWs.Cells(r, 1).Value = keyG
                    destWs.Rows(r).OutlineLevel = 4
                    r = r + 1
                Next keyG
            Next keyF
        Next keyE
    Next keyB
    
    destWs.Outline.ShowLevels RowLevels:=3
    destWs.Outline.ShowLevels RowLevels:=2
    destWs.Outline.ShowLevels RowLevels:=1
    Application.ScreenUpdating = True
    
    MsgBox "Структура успешно создана в новой книге!", vbInformation
End Sub
