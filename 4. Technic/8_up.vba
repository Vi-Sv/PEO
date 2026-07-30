Sub HideZerosAndKeepOrange()
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
    
    Dim totalRng As Range
    Set totalRng = ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, lastCol))
    totalRng.FormatConditions.Delete
    
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
        Set cf = rngSubgroups.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(38, 38, 38)
        cf.SetFirstPriority
    End If
    
    If Not rngMainLines Is Nothing Then
        Set cf = rngMainLines.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(43, 56, 75)
        cf.SetFirstPriority
    End If
    
    If Not rngOthers Is Nothing Then
        Set cf = rngOthers.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        cf.Font.Color = RGB(255, 255, 255)
        cf.SetFirstPriority
    End If
    
    Set cf = totalRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0")
    With cf.Font
        .Color = RGB(228, 108, 10)
        .Bold = True
        .Italic = True
    End With
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
