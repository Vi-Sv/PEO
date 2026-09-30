напишем отдельно: 4. Цвета:

Макрос движется по графе А, и начиная с 9 строки: натыкаясь на
- Уровень 1, Уровень 2, Уровень 3, Уровень 4: идет в графу С, запоминает цветовое форматирование, форматирование цвета шрифта и центрирование и применяет его ко всей соответствующей строке в графах также применяя "все границы" чтобы обернуть ячейки в границы таблиц: D E G H I K L N O P R S T U V, X - AU, AW- BT, BV - CA, CC- CI, CL-CQ, CS-CY, DB-DG, DI-DO,  DR-DW, DY - EE, EH-EM, EO-EU, EX-FC, FE-FK, FO-FR, FT-FY, GA-GD, GF-GX, GZ-HO, HQ-IB, ID-IW, IZ-JR, JT-KI, KK-KV, KX-LO, LR - MD, MF-MS, MU-NH, NJ- NW, NY-OJ, OM-PG, PI-PV, PX-QK, QM-RB, RE-RU, RW-SJ, SL-SY, TA-TN,TP-TW, TZ-UX, UZ-VM, VO-WB, WD-WQ, WT-XL, XN-YA, YC-YP, YR-ZE, ZI-ZQ

Sub ApplyColorsFontsAndBorders()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim arrCols() As String
    Dim targetRange As Range
    Dim srcCell As Range
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Включение режима максимального ускорения Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Форматируем список одиночных столбцов и диапазонов под строгий формат X:X для работы Intersect
    arrCols = Split("D:D,E:E,G:G,H:H,I:I,K:K,L:L,N:N,O:O,P:P,R:R,S:S,T:T,U:U,V:V,X:AU,AW:BT,BV:CA,CC:CI,CL:CQ,CS:CY,DB:DG,DI:DO,DR:DW,DY:EE,EH:EM,EO:EU,EX:FC,FE:FK,FO:FR,FT:FY,GA:GD,GF:GX,GZ:HO,HQ:IB,ID:IW,IZ:JR,JT:KI,KK:KV,KX:LO,LR:MD,MF:MS,MU:NH,NJ:NW,NY:OJ,OM:PG,PI:PV,PX:QK,QM:RB,RE:RU,RW:SJ,SL:SY,TA:TN,TP:TW,TZ:UX,UZ:VM,VO:WB,WD:WQ,WT:XL,XN:YA,YC:YP,YR:ZE,ZI:ZQ", ",")
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        
        ' Проверяем наличие любого из четырех уровней в графе А
        If valA = "Уровень 1" Or valA = "Уровень 2" Or valA = "Уровень 3" Or valA = "Уровень 4" Then
            Set srcCell = ws.Cells(i, "C")
            
            For j = 0 To UBound(arrCols)
                Set targetRange = Intersect(ws.Rows(i), ws.Range(arrCols(j)))
                
                If Not targetRange Is Nothing Then
                    ' 1. Перенос цветового форматирования заливки ячейки
                    targetRange.Interior.Color = srcCell.Interior.Color
                    
                    ' 2. Перенос параметров шрифта (Цвет, Жирность, Курсив, Размер, Шрифт)
                    With targetRange.Font
                        .Color = srcCell.Font.Color
                        .Bold = srcCell.Font.Bold
                        .Italic = srcCell.Font.Italic
                        .Size = srcCell.Font.Size
                        .Name = srcCell.Font.Name
                    End With
                    
                    ' 3. Перенос выравнивания (Горизонтальное и Вертикальное центрирование)
                    targetRange.HorizontalAlignment = srcCell.HorizontalAlignment
                    targetRange.VerticalAlignment = srcCell.VerticalAlignment
                    
                    ' 4. Применение стиля "все границы" (Тонкая сплошная сетка)
                    With targetRange.Borders
                        .LineStyle = xlContinuous
                        .ColorIndex = xlAutomatic
                        .Weight = xlThin
                    End With
                End If
            Next j
        End If
    Next i
    
    ' Восстановление стандартных настроек Excel
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
