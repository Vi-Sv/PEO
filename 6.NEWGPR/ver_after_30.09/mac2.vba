Sub MasterProcessor_Part2A_FormatsAndLevel4()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim colName As String
    
    ' Списки столбцов для пакетного изменения числовых форматов
    Dim arrDec() As String, arrPct() As String, arrL4() As String
    Dim rngsL4 As Variant, rItem As Variant
    Dim cellFormulasL4 As Object
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Нарезка столбцов под форматы из ТЗ
    arrDec = Split("G,H,I,N,O,P,CI,CY,DO,EE,EU,FK", ",")
    arrPct = Split("CA,CH,CQ,CX,DF,DN,DW,ED,EM,ET,FC,FJ,ZL,ZP", ",")
    arrL4 = Split("G,I,BW,BX,BY,BZ,CA,CC,CD,CE,CF,CG,CH,CI", ",")

    ' ==========================================
    ' 6, 7, 8. НОРМАЛИЗАЦИЯ ФОРМАТОВ (ЧИСЛА, ДАТЫ, ПРОЦЕНТЫ)
    ' ==========================================
    For i = 9 To lastRow
        ' Пункт 7: Применение формата дат в грахах K и L
        ws.Cells(i, "K").NumberFormat = "dd.mm.yyyy"
        ws.Cells(i, "L").NumberFormat = "dd.mm.yyyy"
        
        ' Пункт 6: Числовой формат с двумя знаками после запятой (одиночные столбцы)
        For j = 0 To UBound(arrDec)
            ws.Cells(i, arrDec(j)).NumberFormat = "0.00"
        Next j
        
        ' Пункт 6: Числовой формат с двумя знаками после запятой (диапазоны)
        ws.Range(ws.Cells(i, "R"), ws.Cells(i, "V")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "X"), ws.Cells(i, "AU")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "AW"), ws.Cells(i, "BT")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "BV"), ws.Cells(i, "BZ")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "CC"), ws.Cells(i, "CG")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "CL"), ws.Cells(i, "CP")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "CS"), ws.Cells(i, "CW")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "DB"), ws.Cells(i, "DF")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "DI"), ws.Cells(i, "DM")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "DR"), ws.Cells(i, "DV")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "DY"), ws.Cells(i, "EC")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "EH"), ws.Cells(i, "EL")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "EO"), ws.Cells(i, "ES")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "EX"), ws.Cells(i, "FB")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "FE"), ws.Cells(i, "FI")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "FO"), ws.Cells(i, "ZF")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "ZI"), ws.Cells(i, "ZK")).NumberFormat = "0.00"
        ws.Range(ws.Cells(i, "ZM"), ws.Cells(i, "ZO")).NumberFormat = "0.00"
        
        ' Пункт 8: Применение формата процентов (0.00%)
        For j = 0 To UBound(arrPct)
            ws.Cells(i, arrPct(j)).NumberFormat = "0.00%"
        Next j
    Next i

    ' ==========================================
    ' 9. УРОВЕНЬ 4: ТИРАЖИРОВАНИЕ ФОРМУЛ ИЗ СТРОКИ 12
    ' ==========================================
    Set cellFormulasL4 = CreateObject("Scripting.Dictionary")
    
    ' Запоминаем базовые формулы из одиночных столбцов строки 12
    For j = 0 To UBound(arrL4)
        colName = arrL4(j)
        cellFormulasL4(colName) = ws.Cells(12, colName).FormulaR1C1
    Next j
    
    ' Запоминаем формулы из всех диапазонов столбцов строки 12
    rngsL4 = Array("R:V", "X:AU", "AW:BT", "CC:CI", "CL:CQ", "CS:CY", "DB:DG", "DI:DO", "DR:DW", "DY:EE", "EH:EM", "EO:EU", "EX:FC", "FE:FK", "GF:GJ", "IZ:JD", "LR:LV", "OM:OQ", "RE:RI", "TZ:UD", "WT:WX", "ZI:ZQ")
    For Each rItem In rngsL4
        For j = ws.Range(rItem).Column To ws.Range(rItem).Column + ws.Range(rItem).Columns.Count - 1
            cellFormulasL4(ws.Cells(12, j).Address(False, False)) = ws.Cells(12, j).FormulaR1C1
        Next j
    Next rItem
    
    ' Переносим формулы на все строки с признаком "Уровень 4"
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        If valA = "Уровень 4" Then
            For j = 0 To UBound(arrL4)
                colName = arrL4(j)
                ws.Cells(i, colName).FormulaR1C1 = cellFormulasL4(colName)
            Next j
            For Each rItem In rngsL4
                For j = ws.Range(rItem).Column To ws.Range(rItem).Column + ws.Range(rItem).Columns.Count - 1
                    ws.Cells(i, j).FormulaR1C1 = cellFormulasL4(ws.Cells(12, j).Address(False, False))
                Next j
            Next rItem
        End If
    Next i
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
