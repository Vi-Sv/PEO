Sub InsertFormulasToReportWithSmartAnchors()
    Dim wsReport As Worksheet, wsTech As Worksheet
    Dim lastRowReport As Long, lastColReport As Long, lastRowTech As Long
    Dim i As Long, j As Long, offset As Long
    Dim colLetterReport As String
    Dim reportA As Variant, reportB As Variant
    Dim currentParentRow As Long
    Dim cellFormula As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wsReport = ThisWorkbook.Sheets("Отчет")
    Set wsTech = ThisWorkbook.Sheets("Техника")
    
    ' Находим границы листа Отчет
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "B").End(xlUp).Row
    lastColReport = wsReport.UsedRange.Columns.Count
    Do While wsReport.Cells(1, lastColReport).Value = "" And lastColReport > 2
        lastColReport = lastColReport - 1
    Loop
    
    ' Находим последнюю строку на листе Техника
    lastRowTech = wsTech.Cells(wsTech.Rows.Count, "B").End(xlUp).Row
    
    reportA = wsReport.Range("A1:A" & lastRowReport).Value
    reportB = wsReport.Range("B1:B" & lastRowReport).Value
    
    ' 1. Автоматически прописываем формулы-якори во вторую строку отчета
    wsReport.Range(wsReport.Cells(2, 3), wsReport.Cells(2, lastColReport)).NumberFormat = "General"
    For j = 3 To lastColReport
        colLetterReport = Split(wsReport.Cells(1, j).Address, "$")(1)
        wsReport.Cells(2, j).FormulaLocal = "=ПОИСКПОЗ(" & colLetterReport & "$1;Техника!$B$1:$B$" & lastRowTech & ";0)"
    Next j
    
    ' 2. Разносим горизонтальные формулы СУММЕСЛИМН по ячейкам подгрупп
    For j = 3 To lastColReport
        colLetterReport = Split(wsReport.Cells(1, j).Address, "$")(1)
        wsReport.Range(wsReport.Cells(3, j), wsReport.Cells(lastRowReport, j)).NumberFormat = "General"
        
        For i = 3 To lastRowReport
            Dim valA As String: valA = Trim(CStr(reportA(i, 1)))
            
            ' Если строка является подгруппой (М, Д, П, Л, З)
            If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
                If valA = "М" Then offset = 1
                If valA = "Д" Then offset = 2
                If valA = "П" Then offset = 3
                If valA = "Л" Then offset = 4
                If valA = "З" Then offset = 5
                
                currentParentRow = i - offset
                
                ' Сборка текста горизонтальной формулы СУММЕСЛИМН с жесткой границей ABM
                cellFormula = "=СУММЕСЛИМН(ИНДЕКС(Техника!$F$1:$ABM$500;" & colLetterReport & "$2;0);" & _
                              "Техника!$F$4:$ABM$4;""Часы"";" & _
                              "ИНДЕКС(Техника!$G$1:$ABN$500;" & colLetterReport & "$2;0);$A$" & currentParentRow & ";" & _
                              "ИНДЕКС(Техника!$H$1:$ABO$500;" & colLetterReport & "$2;0);""Bound"")"
                
                cellFormula = Replace(cellFormula, """Bound""", """" & valA & """")
                wsReport.Cells(i, j).FormulaLocal = cellFormula
                                 
            ' Если это строка основной числовой группы (1, 2, 3...)
            ElseIf Trim(CStr(reportB(i, 1))) <> "" And IsNumeric(valA) And valA <> "" Then
                wsReport.Cells(i, j).FormulaLocal = "=СУММ(" & colLetterReport & (i + 1) & ":" & colLetterReport & (i + 5) & ")"
            End If
        Next i
    Next j
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
