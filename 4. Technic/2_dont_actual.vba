Sub InsertFormulasToReportScalableFinal()
    Dim wsReport As Worksheet, wsTech As Worksheet
    Dim lastRowReport As Long, lastColReport As Long
    Dim lastRowTech As Long, lastColTech As Long
    Dim i As Long, j As Long, offset As Long
    Dim colLetterReport As String
    Dim colEndTech As String, colShift1Tech As String, colShift2Tech As String
    Dim reportA As Variant, reportB As Variant
    Dim cellFormula As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wsReport = ThisWorkbook.Sheets("Отчет")
    Set wsTech = ThisWorkbook.Sheets("Техника")
    
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "B").End(xlUp).Row
    
    lastColReport = wsReport.UsedRange.Columns.Count
    Do While wsReport.Cells(1, lastColReport).Value = "" And lastColReport > 2
        lastColReport = lastColReport - 1
    Loop
    
    lastRowTech = wsTech.Cells(wsTech.Rows.Count, "B").End(xlUp).Row
    lastColTech = wsTech.Cells(4, wsTech.Columns.Count).End(xlToLeft).Column
    
    colEndTech = Split(wsTech.Cells(1, lastColTech).Address, "$")(1)
    colShift1Tech = Split(wsTech.Cells(1, lastColTech + 1).Address, "$")(1)
    colShift2Tech = Split(wsTech.Cells(1, lastColTech + 2).Address, "$")(1)
    
    reportA = wsReport.Range("A1:A" & lastRowReport).Value
    reportB = wsReport.Range("B1:B" & lastRowReport).Value
    
    For j = 3 To lastColReport
        colLetterReport = Split(wsReport.Cells(1, j).Address, "$")(1)
        
        wsReport.Range(wsReport.Cells(3, j), wsReport.Cells(lastRowReport, j)).NumberFormat = "General"
        
        For i = 3 To lastRowReport
            Dim valA As String: valA = Trim(CStr(reportA(i, 1)))
            
            If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
                If valA = "М" Then offset = 1
                If valA = "Д" Then offset = 2
                If valA = "П" Then offset = 3
                If valA = "Л" Then offset = 4
                If valA = "З" Then offset = 5
                
                cellFormula = "=SUMPRODUCT((Техника!$B$5:$B$" & lastRowTech & "=" & colLetterReport & "$1)*" & _
                              "(Техника!$AC$5:$" & colShift1Tech & "$" & lastRowTech & "=$A" & (i - offset) & ")*" & _
                              "(Техника!$AD$5:$" & colShift2Tech & "$" & lastRowTech & "=$A" & i & ")*" & _
                              "N(Техника!$AB$5:$" & colEndTech & "$" & lastRowTech & "))"
                wsReport.Cells(i, j).Formula = cellFormula
                                 
            ElseIf Trim(CStr(reportB(i, 1))) <> "" And IsNumeric(valA) Then
                cellFormula = "=SUM(" & colLetterReport & (i + 1) & ":" & colLetterReport & (i + 5) & ")"
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
