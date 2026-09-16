Sub CreateHierarchyTreeWithHeader()
    Dim srcWb As Workbook, destWb As Workbook
    Dim srcWs As Worksheet, destWs As Worksheet
    Dim dict As Object, lastRow As Long, i As Long, maxCol As Long
    Dim bVal As String, eVal As String, fVal As String, gVal As String
    Dim keyB As Variant, keyE As Variant, keyF As Variant, keyG As Variant
    Dim r As Long, srcRow As Long, colCount As Long
    Dim keysArr() As Variant, j As Long, temp As Variant
    
    Set srcWb = ActiveWorkbook
    On Error Resume Next
    Set srcWs = srcWb.Sheets("ВР")
    On Error GoTo 0
    
    If srcWs Is Nothing Then
        MsgBox "Ошибка: Лист 'ВР' не найден!", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = srcWs.Cells(srcWs.Rows.Count, "B").End(xlUp).Row
    
    maxCol = srcWs.UsedRange.Columns.Count + srcWs.UsedRange.Column - 1
    If maxCol < 8 Then maxCol = 8
    colCount = maxCol - 8 + 1
    
    For i = 8 To lastRow
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
                            dict(bVal)(eVal)(fVal)(gVal) = i
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    If dict.Count = 0 Then
        MsgBox "Ошибка: Нет данных для обработки.", vbCritical
        Exit Sub
    End If
    
    j = 0
    ReDim keysArr(0 To dict.Count - 1)
    For Each keyB In dict.Keys
        keysArr(j) = keyB
        j = j + 1
    Next keyB
    
    For i = LBound(keysArr) To UBound(keysArr) - 1
        For j = i + 1 To UBound(keysArr)
            If StrComp(CStr(keysArr(i)), CStr(keysArr(j)), vbTextCompare) > 0 Then
                temp = keysArr(i)
                keysArr(i) = keysArr(j)
                keysArr(j) = temp
            End If
        Next j
    Next i
    
    Set destWb = Workbooks.Add
    Set destWs = destWb.Sheets(1)
    destWs.Outline.SummaryRow = xlSummaryAbove
    
    destWs.Cells(1, 1).Value = "Объект"
    destWs.Range(destWs.Cells(1, 2), destWs.Cells(1, colCount + 1)).Value = _
        srcWs.Range(srcWs.Cells(7, 8), srcWs.Cells(7, maxCol)).Value
    
    r = 2
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = LBound(keysArr) To UBound(keysArr)
        keyB = keysArr(i)
        destWs.Cells(r, 1).Value = keyB
        destWs.Rows(r).RowHeight = 30
        destWs.Rows(r).Font.Bold = True
        destWs.Rows(r).OutlineLevel = 1
        r = r + 1
        For Each keyE In dict(keyB).Keys
            destWs.Cells(r, 1).Value = keyE
            destWs.Rows(r).OutlineLevel = 2
            r = r + 1
            For Each keyF In dict(keyB)(keyE).Keys
                destWs.Cells(r, 1).Value = keyF
                destWs.Rows(r).Font.Bold = True
                destWs.Rows(r).Font.Italic = True
                destWs.Rows(r).OutlineLevel = 3
                r = r + 1
                For Each keyG In dict(keyB)(keyE)(keyF).Keys
                    srcRow = dict(keyB)(keyE)(keyF)(keyG)
                    destWs.Cells(r, 1).Value = keyG
                    
                    destWs.Range(destWs.Cells(r, 2), destWs.Cells(r, colCount + 1)).FormulaR1C1 = _
                        srcWs.Range(srcWs.Cells(srcRow, 8), srcWs.Cells(srcRow, maxCol)).FormulaR1C1
                        
                    destWs.Rows(r).OutlineLevel = 4
                    r = r + 1
                Next keyG
            Next keyF
        Next keyE
    Next i
    
    destWs.Outline.ShowLevels RowLevels:=3
    destWs.Outline.ShowLevels RowLevels:=2
    destWs.Outline.ShowLevels RowLevels:=1
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Структура с шапкой успешно создана!", vbInformation
End Sub
