Сейчас на всех уровнях 4 в графах:
ZM формула (=BE)
ZN формула (=JB)
задача такова
написать макрос, который заменит в данных графах формулу равенства на формулу с услвоием:
Если в ячейке ZM2 (графа ZM, строка 2) сентябрь 2026
то в графах будет именно это 
ZM формула (=BE)
ZN формула (=JB)

однако, если "октябрь 2026"
то 
ZM формула (=BF)
ZN формула (=JB)
если "ноябрь 2026"
то 
ZM формула (=BG)
ZN формула (=OO)
если "декабрь 2026"
то 
ZM формула (=BH)
ZN формула (=RG)
если "январь 2027"
то 
ZM формула (=BI)
ZN формула (=UB)
если "февраль 2027"
то 
ZM формула (=BJ)
ZN формула (=WV)



Sub UpdateFormulasZM_ZN_WithFixedCondition()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valA As String
    Dim formulaZM As String
    Dim formulaZN As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' Проверяем, если строка является "Уровень 4"
        If valA = "Уровень 4" Then
            
            ' Сборка формулы ЕСЛИ для графы ZM с абсолютной ссылкой на $ZM$2
            formulaZM = "=IF($ZM$2=""сентябрь 2026"",BE" & i & "," & _
                        "IF($ZM$2=""октябрь 2026"",BF" & i & "," & _
                        "IF($ZM$2=""ноябрь 2026"",BG" & i & "," & _
                        "IF($ZM$2=""декабрь 2026"",BH" & i & "," & _
                        "IF($ZM$2=""январь 2027"",BI" & i & "," & _
                        "IF($ZM$2=""февраль 2027"",BJ" & i & ",BE" & i & "))))))"
            
            ' Сборка формулы ЕСЛИ для графы ZN с абсолютной ссылкой на $ZM$2
            formulaZN = "=IF($ZM$2=""сентябрь 2026"",JB" & i & "," & _
                        "IF($ZM$2=""октябрь 2026"",JB" & i & "," & _
                        "IF($ZM$2=""ноябрь 2026"",OO" & i & "," & _
                        "IF($ZM$2=""декабрь 2026"",RG" & i & "," & _
                        "IF($ZM$2=""январь 2027"",UB" & i & "," & _
                        "IF($ZM$2=""февраль 2027"",WV" & i & ",JB" & i & "))))))"
            
            ' Вставляем готовые формулы в текущую строку
            ws.Cells(i, "ZM").Formula = formulaZM
            ws.Cells(i, "ZN").Formula = formulaZN
        End If
    Next i
    
    ' Восстановление настроек Excel и пересчет данных
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ws.Calculate
End Sub
