напишем макрос отдельно: 10. Формула для объемов:
Макрос движется сверху вниз по графе А, если находит Уровень 3, идет в графу G, и копирует формулу в следующе графы: N R S AW-BT  BV BX CC CE CI CL CN CS CU CY DB DD DI DK DO DR DT DY EA EE EH EJ EO EQ EU EX EZ FE FG FK GF GH IZ JB LR LT OM OO RE RG TZ UB WT WV ZI ZJ ZM ZN ZO
------------------

Sub CopyFormulaLevel3Volumes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim fG As String
    Dim targetCols() As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Массив всех одиночных целевых столбцов из вашего ТЗ для переноса формулы из графы G
    targetCols = Split("N,R,S,BV,BX,CC,CE,CI,CL,CN,CS,CU,CY,DB,DD,DI,DK,DO,DR,DT,DY,EA,EE,EH,EJ,EO,EQ,EU,EX,EZ,FE,FG,FK,GF,GH,IZ,JB,LR,LT,OM,OO,RE,RG,TZ,UB,WT,WV,ZI,ZJ,ZM,ZN,ZO", ",")
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' Проверяем, если строка является "Уровень 3"
        If valA = "Уровень 3" Then
            ' Запоминаем локальную формулу из графы G текущей строки
            fG = ws.Cells(i, "G").FormulaR1C1
            
            ' 1. Копируем формулу в одиночные столбцы из массива
            For j = 0 To UBound(targetCols)
                ws.Cells(i, targetCols(j)).FormulaR1C1 = fG
            Next j
            
            ' 2. Копируем формулу в непрерывный диапазон столбцов AW-BT (включая AW и BT)
            For j = ws.Columns("AW").Column To ws.Columns("BT").Column
                ws.Cells(i, j).FormulaR1C1 = fG
            Next j
        End If
    Next i
    
    ' Восстановление стандартных настроек Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
