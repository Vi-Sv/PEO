Sub HideZerosWithFontColorOptimized()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim valA As String, valB As String
    Dim rngSubgroups As Range, rngMainLines As Range, rngOthers As Range
    Dim cf As FormatCondition
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 3 Or lastCol < 3 Then Exit Sub
    
    For i = 3 To lastRow
        valA = Trim(CStr(ws.Cells(i, "A").Value))
        valB = Trim(CStr(ws.Cells(i, "B").Value))
        
        If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
            If rngSubgroups Is Nothing Then Set rngSubgroups = ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)) Else Set rngSubgroups = Union(rngSubgroups, ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)))
        ElseIf valB <> "" And IsNumeric(valA) Then
            If rngMainLines Is Nothing Then Set rngMainLines = ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)) Else Set rngMainLines = Union(rngMainLines, ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)))
        Else
            If rngOthers Is Nothing Then Set rngOthers = ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)) Else Set rngOthers = Union(rngOthers, ws.Range(ws.Cells(i, 3), ws.Cells(i, lastCol)))
        End If
    Next i
    
    If Not rngSubgroups Is Nothing Then
        rngSubgroups.FormatConditions.Delete
        Set cf = rngSubgroups.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(38, 38, 38)
    End If
    
    If Not rngMainLines Is Nothing Then
        rngMainLines.FormatConditions.Delete
        Set cf = rngMainLines.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(43, 56, 75)
    End If
    
    If Not rngOthers Is Nothing Then
        rngOthers.FormatConditions.Delete
        Set cf = rngOthers.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(255, 255, 255)
    End If
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
