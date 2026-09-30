Sub MasterProcessor_Part2B_Levels3_2_AndGrouping()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim fG As String
    
    ' Списки для переноса формул Уровней 2 и 3
    Dim arrL3() As String, arrL2_1() As String, arrL2_2() As String
    Dim dictL3 As Object, dictL2_1 As Object, dictL2_2 As Object
    
    ' Переменные для группировки верхней панели
    Dim arrHeader() As String, headerItem As Variant
    Dim g1 As Variant, g2 As Variant
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Инициализация массивов для формул на основе ТЗ
    arrL3 = Split("CA,CH,CI,CQ,CX,CY,DG,DN,DO,DW,ED,EE,EM,ET,EU,FC,FJ,FK,ZL,ZP,ZQ", ",")
    arrL2_1 = Split("CA,CH,CQ,CX,DG,DN,DW,ED,EM,ET,FC,FJ,ZL,ZM", ",")
    ' Пункт 13: Исправлено - вместо опечатки "I" теперь строго стоит "DI"
    arrL2_2 = Split("BV,BX,CC,CE,CH,CI,CL,CN,CQ,CS,CU,CX,DB,DD,DF,DG,DI,DK,DN,DR,DT,DW,DY,EA,ED,EK,EJ,EH,EO,EQ,ES,ET,EU,EX,EZ,FA,FC,FE,FG,FH,FJ,FK", ",")
    
    Set dictL3 = CreateObject("Scripting.Dictionary")
    Set dictL2_1 = CreateObject("Scripting.Dictionary")
    Set dictL2_2 = CreateObject("Scripting.Dictionary")
    
    ' ==========================================
    ' ЗАПОМИНАНИЕ ИСХОДНЫХ ДАННЫХ ИЗ СТРОК 10 И 11
    ' ==========================================
    ' Пункт 12: Запоминаем формулы Уровня 3 (Строка 11)
    For j = 0 To UBound(arrL3)
        dictL3(arrL3(j)) = ws.Cells(11, arrL3(j)).FormulaR1C1
    Next j
    
    ' Пункт 12: Запоминаем формулы Уровня 2 (Строка 10)
    For j = 0 To UBound(arrL2_1)
        dictL2_1(arrL2_1(j)) = ws.Cells(10, arrL2_1(j)).FormulaR1C1
    Next j
    
    ' Пункт 13: Запоминаем формулы/пустоту Уровня 2 (Строка 10)
    For j = 0 To UBound(arrL2_2)
        If ws.Cells(10, arrL2_2(j)).HasFormula Then
            dictL2_2(arrL2_2(j)) = ws.Cells(10, arrL2_2(j)).FormulaR1C1
        Else
            dictL2_2(arrL2_2(j)) = ws.Cells(10, arrL2_2(j)).Value
        End If
    Next j

    ' ==========================================
    ' СКВОЗНОЙ ЦИКЛ ОБРАБОТКИ УРОВНЕЙ 3 И 2
    ' ==========================================
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' ------------------------------------------
        ' 10. ПРИМЕНЕНИЕ ФОРМУЛ ОБЪЕМОВ ДЛЯ УРОВНЯ 3
        ' ------------------------------------------
        If valA = "Уровень 3" Then
            fG = ws.Cells(i, "G").FormulaR1C1
            ws.Cells(i, "N").FormulaR1C1 = fG: ws.Cells(i, "R").FormulaR1C1 = fG: ws.Cells(i, "S").FormulaR1C1 = fG
            ws.Cells(i, "GF").FormulaR1C1 = fG: ws.Cells(i, "GH").FormulaR1C1 = fG: ws.Cells(i, "IZ").FormulaR1C1 = fG
            ws.Cells(i, "JB").FormulaR1C1 = fG: ws.Cells(i, "LR").FormulaR1C1 = fG: ws.Cells(i, "LT").FormulaR1C1 = fG
            ws.Cells(i, "OM").FormulaR1C1 = fG: ws.Cells(i, "OO").FormulaR1C1 = fG: ws.Cells(i, "RE").FormulaR1C1 = fG
            ws.Cells(i, "RG").FormulaR1C1 = fG: ws.Cells(i, "TZ").FormulaR1C1 = fG: ws.Cells(i, "UB").FormulaR1C1 = fG
            ws.Cells(i, "WT").FormulaR1C1 = fG: ws.Cells(i, "WV").FormulaR1C1 = fG: ws.Cells(i, "ZI").FormulaR1C1 = fG
            ws.Cells(i, "ZJ").FormulaR1C1 = fG: ws.Cells(i, "ZM").FormulaR1C1 = fG: ws.Cells(i, "ZN").FormulaR1C1 = fG
            ws.Cells(i, "ZO").FormulaR1C1 = fG
            
            ' Заполнение формулой G диапазона AW-BT
            For j = ws.Columns("AW").Column To ws.Columns("BT").Column
                ws.Cells(i, j).FormulaR1C1 = fG
            Next j
            
            ' Накатывание особых формул Уровня 3 (Пункт 12)
            For j = 0 To UBound(arrL3)
                ws.Cells(i, arrL3(j)).FormulaR1C1 = dictL3(arrL3(j))
            Next j
            
        ' ------------------------------------------
        ' 12 и 13. ПРИМЕНЕНИЕ ФОРМУЛ ДЛЯ УРОВНЯ 2
        ' ------------------------------------------
        ElseIf valA = "Уровень 2" Then
            ' Особые формулы Уровня 2 (Пункт 12)
            For j = 0 To UBound(arrL2_1)
                ws.Cells(i, arrL2_1(j)).FormulaR1C1 = dictL2_1(arrL2_1(j))
            Next j
            
            ' Формулы объемов 15-30 Уровня 2 (Пункт 13)
            For j = 0 To UBound(arrL2_2)
                If Left(CStr(dictL2_2(arrL2_2(j))), 1) = "=" Then
                    ws.Cells(i, arrL2_2(j)).FormulaR1C1 = dictL2_2(arrL2_2(j))
                Else
                    ws.Cells(i, arrL2_2(j)).Value = dictL2_2(arrL2_2(j))
                End If
            Next j
        End If
    Next i

    ' ==========================================
    ' 11. ВЕРХНЯЯ ПАНЕЛЬ: ТРЕХУРОВНЕВАЯ ВЛОЖЕННАЯ ГРУППИРОВКА
    ' Собрано строго по матрешке от крупных к мелким блокам
    ' ==========================================
    On Error Resume Next
    ' Разгруппировываем только столбцы, чтобы не задеть плюсики строк
    ws.Columns.Ungroup
    
    ' --- Уровень 1: Главные родительские диапазоны верхней панели ---
    For Each g1 In Array("C:E", "G:I", "K:L", "N:P", "R:V", "X:AU", "AW:BT", "BV:CJ", "CL:CZ", "DB:DP", "DR:EF", "EH:EV", "EX:FL", "FO:FR", "FT:FY", "GA:GD", "GF:IX", "IZ:LP", "LR:OK", "OM:RC", "RE:TX", "TZ:WR", "WT:ZF", "ZI:ZQ")
        ws.Columns(g1).Group
    Next g1
    
    ' --- Уровень 2: Вложенные подгруппы первого порядка ---
    For Each g2 In Array("BV:CA", "CC:CI", "CL:CQ", "CS:CY", "DB:DG", "DI:DO", "DR:DW", "DY:EE", "EH:EM", "EO:EU", "EX:FC", "FE:FK", "GF:GX", "GZ:HO", "HQ:IB", "ID:IW", "IZ:JR", "JT:KI", "KK:KV", "KX:LO", "LR:MD", "MF:MS", "MU:NH", "NJ:NW", "NY:OK", "OM:PG", "PI:PV", "PX:QK", "QM:RB", "RE:RU", "RW:SJ", "SL:SY", "TA:TN", "TP:TW", "TZ:UX", "UZ:VM", "VO:WB", "WD:WQ", "WT:XL", "XN:YA", "YC:YP", "YR:ZF")
        ws.Columns(g2).Group
    Next g2
    
    ' --- Уровень 3: Самые глубокие внутренние группы второго порядка ---
    ' Создают точечные "батарейки" внутри групп Уровня 2
    ws.Columns("BX:BY").Group
    ws.Columns("CD:CE").Group
    ws.Columns("CM:CN").Group
    ws.Columns("CT:CU").Group
    ws.Columns("DD:DE").Group
    ws.Columns("DK:DL").Group
    ws.Columns("DT:DU").Group
    ws.Columns("DZ:EA").Group
    ws.Columns("EJ:EK").Group
    ws.Columns("ER:ES").Group
    ws.Columns("FA:FB").Group
    ws.Columns("FH:FI").Group
    On Error GoTo 0
    
    ' Восстановление стандартного окружения Excel
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
