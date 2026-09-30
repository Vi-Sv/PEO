Sub MasterProcessor_Part1_StructureColorsAndLabor()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim valA As String
    Dim arrCols() As String
    Dim fI As String
    Dim targetRange As Range
    Dim arrTargetI() As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    
    ' Включение максимального ускорения VBA
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ==========================================
    ' 1. НУМЕРАЦИЯ (С 9 строки по конец листа)
    ' ==========================================
    Dim c1 As Long: c1 = 0
    Dim c2 As Long: c2 = 0
    Dim c3 As Long: c3 = 0
    Dim c4 As Long: c4 = 0
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        Select Case valA
            Case "Уровень 1"
                c1 = c1 + 1: c2 = 0: c3 = 0: c4 = 0
                ws.Cells(i, "C").Value = "'" & c1
            Case "Уровень 2"
                If c1 > 0 Then
                    c2 = c2 + 1: c3 = 0: c4 = 0
                    ws.Cells(i, "C").Value = "'" & c1 & "." & c2
                End If
            Case "Уровень 3"
                If c1 > 0 And c2 > 0 Then
                    c3 = c3 + 1: c4 = 0
                    ws.Cells(i, "C").Value = "'" & c1 & "." & c2 & "." & c3
                End If
            Case "Уровень 4"
                If c1 > 0 And c2 > 0 And c3 > 0 Then
                    c4 = c4 + 1
                    ws.Cells(i, "C").Value = "'" & c1 & "." & c2 & "." & c3 & "." & c4
                End If
        End Select
    Next i

    ' ==========================================
    ' 2. ПЛЮСЫ ПО СТРОКАМ (ОЧИСТКА И ГРУППИРОВКА)
    ' ==========================================
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=8
    ws.Cells.ClearOutline
    On Error GoTo 0
    ws.Outline.SummaryRow = xlSummaryAbove
    
    Dim currentL1 As Long: currentL1 = 0
    Dim currentL2 As Long: currentL2 = 0
    Dim currentL3 As Long: currentL3 = 0
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        Select Case valA
            Case "Уровень 1"
                If currentL3 > 0 Then ws.Range(ws.Rows(currentL3 + 1), ws.Rows(i - 1)).Group
                If currentL2 > 0 Then ws.Range(ws.Rows(currentL2 + 1), ws.Rows(i - 1)).Group
                If currentL1 > 0 Then ws.Range(ws.Rows(currentL1 + 1), ws.Rows(i - 1)).Group
                currentL1 = i: currentL2 = 0: currentL3 = 0
            Case "Уровень 2"
                If currentL3 > 0 Then ws.Range(ws.Rows(currentL3 + 1), ws.Rows(i - 1)).Group
                If currentL2 > 0 Then ws.Range(ws.Rows(currentL2 + 1), ws.Rows(i - 1)).Group
                currentL2 = i: currentL3 = 0
            Case "Уровень 3"
                If currentL3 > 0 Then ws.Range(ws.Rows(currentL3 + 1), ws.Rows(i - 1)).Group
                currentL3 = i
        End Select
    Next i
    
    If currentL3 > 0 And lastRow > currentL3 Then ws.Range(ws.Rows(currentL3 + 1), ws.Rows(lastRow)).Group
    If currentL2 > 0 And lastRow > currentL2 Then ws.Range(ws.Rows(currentL2 + 1), ws.Rows(lastRow)).Group
    If currentL1 > 0 And lastRow > currentL1 Then ws.Range(ws.Rows(currentL1 + 1), ws.Rows(lastRow)).Group

    ' ==========================================
    ' 4 (Первая). ФОРМАТИРОВАНИЕ НУМЕРАЦИИ (ГРАФА C)
    ' ==========================================
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        Select Case valA
            Case "Уровень 1": ws.Cells(9, "C").Copy: ws.Cells(i, "C").PasteSpecial Paste:=xlPasteFormats
            Case "Уровень 2": ws.Cells(10, "C").Copy: ws.Cells(i, "C").PasteSpecial Paste:=xlPasteFormats
            Case "Уровень 3": ws.Cells(11, "C").Copy: ws.Cells(i, "C").PasteSpecial Paste:=xlPasteFormats
            Case "Уровень 4": ws.Cells(12, "C").Copy: ws.Cells(i, "C").PasteSpecial Paste:=xlPasteFormats
        End Select
    Next i

    ' ==========================================
    ' 4 (Вторая). ЦВЕТА И СЕТКА ГРАНИЦ ПО ВСЕМ СТРОКАМ
    ' ==========================================
    arrCols = Split("D:D,E:E,G:G,H:H,I:I,K:K,L:L,N:N,O:O,P:P,R:R,S:S,T:T,U:U,V:V,X:AU,AW:BT,BV:CA,CC:CI,CL:CQ,CS:CY,DB:DG,DI:DO,DR:DW,DY:EE,EH:EM,EO:EU,EX:FC,FE:FK,FO:FR,FT:FY,GA:GD,GF:GX,GZ:HO,HQ:IB,ID:IW,IZ:JR,JT:KI,KK:KV,KX:LO,LR:MD,MF:MS,MU:NH,NJ:NW,NY:OJ,OM:PG,PI:PV,PX:QK,QM:RB,RE:RU,RW:SJ,SL:SY,TA:TN,TP:TW,TZ:UX,UZ:VM,VO:WB,WD:WQ,WT:XL,XN:YA,YC:YP,YR:ZE,ZI:ZQ", ",")
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        If valA Like "Уровень [1-4]" Then
            For j = 0 To UBound(arrCols)
                Set targetRange = Intersect(ws.Rows(i), ws.Range(arrCols(j)))
                If Not targetRange Is Nothing Then
                    targetRange.Interior.Color = ws.Cells(i, "C").Interior.Color
                    targetRange.Borders.LineStyle = xlContinuous
                    targetRange.Borders.Weight = xlThin
                End If
            Next j
        End If
    Next i

    ' ==========================================
    ' 5. ФОРМУЛА ДЛЯ ТРУДОЗАТРАТ ДЛЯ УРОВНЕЙ 1-3
    ' ==========================================
    arrTargetI = Split("T,U,V,BW,BY,BZ,CD,CF,CG,CM,CO,CP,CT,CV,CW,DC,DE,DF,DJ,DL,DM,DS,DU,DV,DZ,EB,EC,EI,EK,EL,EP,ER,ES,EY,FA,FB,FF,FH,FI,FP,FR,FU,FW,FY,GB,GD,GG,GI,GJ,JA,JC,JD,LS,LU,LV,ON,OP,OQ,RF,RH,RI,UA,UC,UD,WU,WW,WX", ",")
    
    For i = 9 To lastRow
        valA = Trim(ws.Cells(i, "A").Value)
        If valA = "Уровень 1" Or valA = "Уровень 2" Or valA = "Уровень 3" Then
            fI = ws.Cells(i, "I").FormulaR1C1
            
            For j = 0 To UBound(arrTargetI)
                ws.Cells(i, arrTargetI(j)).FormulaR1C1 = fI
            Next j
            
            For j = ws.Columns("X").Column To ws.Columns("AU").Column
                ws.Cells(i, j).FormulaR1C1 = fI
            Next j
        End If
    Next i

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

