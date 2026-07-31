Sub DiagnoseReportSheet()
    Dim wsReport As Worksheet
    Dim lastRowReport As Long, lastColReport As Long
    Dim i As Long, sampleA As String, sampleB As String
    
    Set wsReport = ThisWorkbook.Sheets("Отчет")
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "B").End(xlUp).Row
    lastColReport = wsReport.Cells(3, wsReport.Columns.Count).End(xlToLeft).Column
    
    Debug.Print "--- ДИАГНОСТИКА ЛИСТА ОТЧЕТ ---"
    Debug.Print "Последняя строка (lastRowReport): " & lastRowReport
    Debug.Print "Последний столбец (lastColReport): " & lastColReport
    
    If lastRowReport >= 3 Then
        Debug.Print "Первые 5 анализируемых строк (Столбец А | Столбец B):"
        For i = 3 To IIf(lastRowReport > 8, 8, lastRowReport)
            sampleA = wsReport.Cells(i, "A").Value
            sampleB = wsReport.Cells(i, "B").Value
            Debug.Print "Строка " & i & ": А='" & sampleA & "' (Длина: " & Len(sampleA) & "); B='" & sampleB & "'"
        Next i
    Else
        Debug.Print "ОШИБКА: lastRowReport меньше 3. Данные не найдены."
    End If
    Debug.Print "-------------------------------"
    MsgBox "Диагностика завершена. Нажмите Ctrl+G в редакторе VBA и скопируйте текст из окна Immediate.", vbInformation
End Sub
