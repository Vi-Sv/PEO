Sub InsertFormulasToReportScalable()
    Dim wsReport As Worksheet, wsTech As Worksheet
    Dim lastRowReport As Long, lastColReport As Long
    Dim lastRowTech As Long, lastColTech As Long
    Dim i As Long, j As Long, offset As Long
    Dim colLetterTech As String, colLetterTechEnd As String
    Dim colLetterTechShift1 As String, colLetterTechShift2 As String
    Dim reportA As Variant, reportB As Variant
    Dim cellFormula As String
    Dim colFormulas() As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wsReport = ThisWorkbook.Sheets("Отчет")
    Set wsTech = ThisWorkbook.Sheets("Техника")
    
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "B").End(xlUp).Row
    lastColReport = wsReport.Cells(1, wsReport.Columns.Count).End(xlToLeft).Column
    
    lastRowTech = wsTech.Cells(wsTech.Rows.Count, "B").End(xlUp).Row
    lastColTech = wsTech.Cells(1, wsTech.Columns.Count).End(xlToLeft).Column
    
    colLetterTech = Split(wsTech.Cells(1, 8).Address, "$")(1) ' "H"
    colLetterTechEnd = Split(wsTech.Cells(1, lastColTech).Address, "$")(1)
    colLetterTechShift1 = Split(wsTech.Cells(1, 9).Address, "$")(1) ' "I"
    colLetterTechShift2 = Split(wsTech.Cells(1, 10).Address, "$")(1) ' "J"
    
    reportA = wsReport.Range("A1:A" & lastRowReport).Value
    reportB = wsReport.Range("B1:B" & lastRowReport).Value
    
    For j = 3 To lastColReport
        Dim colLetterReport As String
        colLetterReport = Split(wsReport.Cells(1, j).Address, "$")(1)
        
        wsReport.Range(wsReport.Cells(3, j), wsReport.Cells(lastRowReport, j)).NumberFormat = "General"
        ReDim colFormulas(3 To lastRowReport, 1 To 1)
        
        For i = 3 To lastRowReport
            Dim valA As String: valA = Trim(CStr(reportA(i, 1)))
            
            If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
                If valA = "М" Then offset = 1
                If valA = "Д" Then offset = 2
                If valA = "П" Then offset = 3
                If valA = "Л" Then offset = 4
                If valA = "З" Then offset = 5
                
                colFormulas(i, 1) = "=SUMPRODUCT((Техника!$" & colLetterTech & "$5:$" & colLetterTechEnd & "$5=" & colLetterReport & "$1)*" & _
                                    "(Техника!$" & colLetterTechShift1 & "$5:$" & Split(wsTech.Cells(1, lastColTech + 1).Address, "$")(1) & "$" & lastRowTech & "=$A" & (i - offset) & ")*" & _
                                    "(Техника!$" & colLetterTechShift2 & "$5:$" & Split(wsTech.Cells(1, lastColTech + 2).Address, "$")(1) & "$" & lastRowTech & "=$A" & i & ")*" & _
                                    "N(Техника!$" & colLetterTech & "$5:$" & colLetterTechEnd & "$" & lastRowTech & "))"
                                 
            ElseIf Trim(CStr(reportB(i, 1))) <> "" And IsNumeric(valA) Then
                colFormulas(i, 1) = "=SUM(" & colLetterReport & (i + 1) & ":" & colLetterReport & (i + 5) & ")"
            End If
        Next i
        
        wsReport.Range(wsReport.Cells(3, j), wsReport.Cells(lastRowReport, j)).Formula = colFormulas
    Next j
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
