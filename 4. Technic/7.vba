Sub ApplyConditionalFormattingRgb()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim rng As Range, cf As FormatCondition
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 3 Or lastCol < 3 Then Exit Sub
    Set rng = ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, lastCol))
    
    rng.FormatConditions.Delete
    
    Set cf = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0")
    With cf.Font
        .Color = RGB(228, 108, 10)
        .Bold = True
    End With
End Sub
