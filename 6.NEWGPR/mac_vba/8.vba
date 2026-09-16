Sub CreateHierarchyTreeWithHeader()
    Dim srcWb As Workbook, destWb As Workbook
    Dim srcWs As Worksheet, destWs As Worksheet
    Dim dict As Object, lastRow As Long, i As Long, maxCol As Long
    Dim bVal As String, eVal As String, fVal As String, gVal As String
    Dim keyB As Variant, keyE As Variant, keyF As Variant, keyG As Variant
    Dim r As Long, srcRow As Long, colCount As Long
    Dim keysArr() As Variant, j As Long, temp As Variant
    
    Dim idx1 As Long, idx2 As Long, idx3 As Long, idx4 As Long
    Dim startR1 As Long, startR2 As Long, startR3 As Long
    
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
            If StrComp(CStr(keysArr(i)), CStr(keysArr(j)), vbTextCompare) < 0 Then
                temp = keysArr(i)
                keysArr(i) = keysArr(j)
                keysArr(j) = temp
            End If
        Next j
    Next i
    
    Set destWb = Workbooks.Add
    Set destWs = destWb.Sheets(1)
    destWs.Outline.SummaryRow = xlSummaryAbove
    
    destWs.Columns("A:A").Insert Shift:=xlToRight
    destWs.Columns("A:A").NumberFormat = "@"
    
    With destWs.Range(destWs.Cells(1, 1), destWs.Cells(1, colCount + 2))
        destWs.Cells(1, 1).Value = "№"
        destWs.Cells(1, 2).Value = "Объект"
        destWs.Range(destWs.Cells(1, 3), destWs.Cells(1, colCount + 2)).Value = _
            srcWs.Range(srcWs.Cells(7, 8), srcWs.Cells(7, maxCol)).Value
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(52, 73, 94)
        .RowHeight = 25
        .VerticalAlignment = xlCenter
    End With
    
    r = 2
    idx1 = 0: idx2 = 0: idx3 = 0: idx4 = 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = LBound(keysArr) To UBound(keysArr)
        keyB = keysArr(i)
        idx1 = idx1 + 1
        startR1 = r
        
        destWs.Cells(r, 1).Value = CStr(idx1)
        destWs.Cells(r, 1).HorizontalAlignment = xlLeft
        destWs.Cells(r, 2).Value = keyB
        With destWs.Rows(r)
            .RowHeight = 30
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(52, 73, 94)
            .VerticalAlignment = xlCenter
            .OutlineLevel = 1
        End With
        r = r + 1
        
        idx2 = 0
        For Each keyE In dict(keyB).Keys
            idx2 = idx2 + 1
            startR2 = r
            
            destWs.Cells(r, 1).Value = CStr(idx1 & "." & idx2)
            destWs.Cells(r, 1).HorizontalAlignment = xlRight
            destWs.Cells(r, 2).Value = keyE
            With destWs.Rows(r)
                .RowHeight = 30
                .Font.Bold = True
                .Font.Color = RGB(44, 62, 80)
                .Interior.Color = RGB(214, 234, 248)
                .VerticalAlignment = xlCenter
                .OutlineLevel = 2
            End With
            r = r + 1
            
            idx3 = 0
            For Each keyF In dict(keyB)(keyE).Keys
                idx3 = idx3 + 1
                startR3 = r
                
                destWs.Cells(r, 1).Value = CStr(idx1 & "." & idx2 & "." & idx3)
                destWs.Cells(r, 1).HorizontalAlignment = xlLeft
                destWs.Cells(r, 2).Value = keyF
                With destWs.Rows(r)
                    .RowHeight = 45
                    .Font.Bold = True
                    .Font.Italic = True
                    .Font.Color = RGB(44, 62, 80)
                    .Interior.Color = RGB(235, 245, 251)
                    .VerticalAlignment = xlCenter
                    .OutlineLevel = 3
                End With
                r = r + 1
                
                idx4 = 0
                For Each keyG In dict(keyB)(keyE)(keyF).Keys
                    srcRow = dict(keyB)(keyE)(keyF)(keyG)
                    idx4 = idx4 + 1
                    destWs.Cells(r, 1).Value = CStr(idx1 & "." & idx2 & "." & idx3 & "." & idx4)
                    destWs.Cells(r, 1).HorizontalAlignment = xlRight
                    destWs.Cells(r, 2).Value = keyG
                    
                    destWs.Range(destWs.Cells(r, 3), destWs.Cells(r, colCount + 2)).FormulaR1C1 = _
                        srcWs.Range(srcWs.Cells(srcRow, 8), srcWs.Cells(srcRow, maxCol)).FormulaR1C1
                        
                    destWs.Cells(r, 7).FormulaR1C1 = "=RC6*RC4"
                    destWs.Cells(r, 7).NumberFormat = "#,##0.00"
                        
                    With destWs.Rows(r)
                        .RowHeight = 45
                        .Font.Bold = False
                        .Font.Italic = False
                        .Font.Color = RGB(0, 0, 0)
                        .Interior.Color = RGB(255, 255, 255)
                        .VerticalAlignment = xlCenter
                        .OutlineLevel = 4
                    End With
                    r = r + 1
                Next keyG
                
                If r - 1 > startR3 Then
                    destWs.Cells(startR3, 7).FormulaR1C1 = "=SUM(R[" & (startR3 + 1 - startR3) & "]C:R[" & (r - 1 - startR3) & "]C)"
                End If
            Next keyF
            
            If r - 1 > startR2 Then
                destWs.Cells(startR2, 7).FormulaR1C1 = "=SUMIF(R[" & (startR2 + 1 - startR2) & "]C:R[" & (r - 1 - startR2) & "]C,""<>"",R[" & (startR2 + 1 - startR2) & "]C:R[" & (r - 1 - startR2) & "]C)/2"
            End If
        Next keyE
        
        If r - 1 > startR1 Then
            destWs.Cells(startR1, 7).FormulaR1C1 = "=SUMIF(R[" & (startR1 + 1 - startR1) & "]C:R[" & (r - 1 - startR1) & "]C,""<>"",R[" & (startR1 + 1 - startR1) & "]C:R[" & (r - 1 - startR1) & "]C)/3"
        End If
    Next i
    
    destWs.Columns("G:G").NumberFormat = "#,##0.00"
    destWs.Columns.AutoFit
    
    destWs.Range("C1,H1,M1,S1,V1,AU1").EntireColumn.ColumnWidth = 4
    
    destWs.Outline.ShowLevels RowLevels:=3
    destWs.Outline.ShowLevels RowLevels:=2
    destWs.Outline.ShowLevels RowLevels:=1
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Структура успешно создана, ширина столбцов настроена!", vbInformation
End Sub
