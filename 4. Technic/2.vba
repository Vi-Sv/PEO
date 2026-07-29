Sub InsertFormulasToReportNew()
    Dim wsReport As Worksheet, lastRowReport As Long, lastColReport As Long
    Dim i As Long, j As Long, offset As Long
    Dim colLetterTech As String
    Dim reportA As Variant, reportB As Variant
    Dim cellFormula As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wsReport = ThisWorkbook.Sheets("Отчет")
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "B").End(xlUp).Row
    lastColReport = wsReport.Cells(1, wsReport.Columns.Count).End(xlToLeft).Column
    
    reportA = wsReport.Range("A1:A" & lastRowReport).Value
    reportB = wsReport.Range("B1:B" & lastRowReport).Value
    
    For j = 3 To lastColReport
        colLetterTech = Split(wsReport.Cells(1, j).Address, "$")(1)
        
        wsReport.Range(wsReport.Cells(3, j), wsReport.Cells(lastRowReport, j)).NumberFormat = "General"
        
        For i = 3 To lastRowReport
            Dim valA As String: valA = Trim(CStr(reportA(i, 1)))
            
            If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
                If valA = "М" Then offset = 1
                If valA = "Д" Then offset = 2
                If valA = "П" Then offset = 3
                If valA = "Л" Then offset = 4
                If valA = "З" Then offset = 5
                
                cellFormula = "=SUMPRODUCT((Техника!$B$5:$B$19=" & colLetterTech & "$1)*" & _
                              "(Техника!$AC$5:$BIA$19=$A" & (i - offset) & ")*" & _
                              "(Техника!$AD$5:$BIB$19=$A" & i & ")*" & _
                              "N(Техника!$AB$5:$BHZ$19))"
                wsReport.Cells(i, j).Formula = cellFormula
                                 
            ElseIf Trim(CStr(reportB(i, 1))) <> "" And IsNumeric(valA) Then
                cellFormula = "=SUM(" & colLetterTech & (i + 1) & ":" & colLetterTech & (i + 5) & ")"
                wsReport.Cells(i, j).Formula = cellFormula
            End If
        Next i
    Next j
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
