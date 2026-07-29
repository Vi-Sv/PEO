Sub FixRowHeightAndAlignment()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim valA As String
    
    Set ws = ThisWorkbook.Sheets("Отчет")
    ws.Rows(2).RowHeight = 55
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 3 To lastRow
        valA = Trim(CStr(ws.Cells(i, "A").Value))
        If valA = "М" Or valA = "Д" Or valA = "П" Or valA = "Л" Or valA = "З" Then
            ws.Cells(i, "B").HorizontalAlignment = xlLeft
        End If
    Next i
End Sub
