Sub ApplyAllGESNNorms()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rName As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 2 To lastRow
        rName = WorksheetFunction.Trim(ws.Cells(i, 2).Value)
        
        Select Case True
            ' --- ТРАНСПОРТ И ЛОГИСТИКА / ПЕРЕВОЗКА И ВЫВОЗ МУСОРА/ГРУНТА (Нормы на 1 т груза) ---
            Case InStr(1, rName, "Перевозка", vbTextCompare) > 0 Or InStr(1, rName, "Вывоз", vbTextCompare) > 0 Or InStr(1, rName, "Транспорт (почвенно-растительный", vbTextCompare) > 0 Or InStr(1, rName, "Транспортировка мусора", vbTextCompare) > 0 Or InStr(1, rName, "Отвозка металлолома", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "т"
                If InStr(1, rName, "30 км", vbTextCompare) > 0 Or InStr(1, rName, "34 км", vbTextCompare) > 0 Or InStr(1, rName, "35 км", vbTextCompare) > 0 Or InStr(1, rName, "32 км", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.095 ' ФССЦпг-03-01-01-030
                ElseIf InStr(1, rName, "7 км", vbTextCompare) > 0 Or InStr(1, rName, "5 км", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.042
                ElseIf InStr(1, rName, "11 км", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.056
                ElseIf InStr(1, rName, "23 км", vbTextCompare) > 0 Or InStr(1, rName, "25 км", vbTextCompare) > 0 Or InStr(1, rName, "27 км", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.081
                ElseIf InStr(1, rName, "1 км", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.021
                Else
                    ws.Cells(i, 5).Value = 0.065
                End If

            ' --- ЗЕМЛЯНЫЕ РАБОТЫ: РАЗРАБОТКА ЭКСКАВАТОРАМИ (Приведение к 1 м3) ---
            Case InStr(1, rName, "Разработка грунта", vbTextCompare) > 0 Or InStr(1, rName, "Разработка траншей", vbTextCompare) > 0 Or InStr(1, rName, "Разработка котлована", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                If InStr(1, rName, "ковшом вместимостью 0,4", vbTextCompare) > 0 Or InStr(1, rName, "0,25 м3", vbTextCompare) > 0 Then
                    If InStr(1, rName, "в отвал", vbTextCompare) > 0 Then
                        ws.Cells(i, 5).Value = 0.0255 ' ГЭСН01-01-013-03
                    Else
                        ws.Cells(i, 5).Value = 0.0312 ' ГЭСН01-01-014-02
                    End If
                ElseIf InStr(1, rName, "ковшом вместимостью 0,65", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.0185 ' ГЭСН01-01-014-03
                ElseIf InStr(1, rName, "ковшом вместимостью 1", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.0124 ' ГЭСН01-01-014-04
                ElseIf InStr(1, rName, "вручную", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 1.62 ' ГЭСН01-02-057-02
                Else
                    ws.Cells(i, 5).Value = 0.022
                End If

            ' --- ЗЕМЛЯНЫЕ РАБОТЫ: ОБРАТНАЯ ЗАСЫПКА И ПЕРЕМЕЩЕНИЕ БУЛЬДОЗЕРАМИ (Приведение к 1 м3) ---
            Case InStr(1, rName, "Обратная засыпка", vbTextCompare) > 0 Or InStr(1, rName, "Засыпка траншей", vbTextCompare) > 0 Or InStr(1, rName, "Засыпка грунта", vbTextCompare) > 0 Or InStr(1, rName, "Перемещение разработанного", vbTextCompare) > 0 Or InStr(1, rName, "Перемещение грунта", vbTextCompare) > 0 Or InStr(1, rName, "Срезка грунта", vbTextCompare) > 0 Or InStr(1, rName, "Срезка почвенно-растительного", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                If InStr(1, rName, "добавлять", vbTextCompare) > 0 Or InStr(1, rName, "Дополнительное перемещение", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.00045 ' ГЭСН01-01-034-05
                Else
                    ws.Cells(i, 5).Value = 0.0016 ' ГЭСН01-01-034-02
                End If

            ' --- РАБОТА НА ОТВАЛЕ (Приведение к 1 м3) ---
            Case InStr(1, rName, "Работа на отвале", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                ws.Cells(i, 5).Value = 0.0041 ' ГЭСН01-01-017-02

            ' --- ПЛАНИРОВКА И УКРЕПЛЕНИЕ ОТКОСОВ / ПОСЕВ ТРАВ ---
            Case InStr(1, rName, "Планировка откосов", vbTextCompare) > 0 Or InStr(1, rName, "Планировка площадей", vbTextCompare) > 0 Or InStr(1, rName, "Планировка площадки", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м2"
                ws.Cells(i, 5).Value = 0.0028 ' ГЭСН01-01-045-01
            Case InStr(1, rName, "Укрепление откосов", vbTextCompare) > 0 Or InStr(1, rName, "Посев многолетних", vbTextCompare) > 0
                If InStr(1, rName, "га", vbTextCompare) > 0 Then
                    ws.Cells(i, 4).Value = "га"
                    ws.Cells(i, 5).Value = 42.5 ' ГЭСН01-02-041-01
                Else
                    ws.Cells(i, 4).Value = "м2"
                    ws.Cells(i, 5).Value = 0.0142 ' ГЭСН01-02-041-01
                End If

            ' --- УПЛОТНЕНИЕ ГРУНТА КАТКАМИ (Приведение к 1 м3) ---
            Case InStr(1, rName, "Уплотнение грунта", vbTextCompare) > 0 Or InStr(1, rName, "Уплотнение насыпи", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                If InStr(1, rName, "последующий проход", vbTextCompare) > 0 Then
                    ws.Cells(i, 5).Value = 0.00021 ' ГЭСН01-02-001-02
                Else
                    ws.Cells(i, 5).Value = 0.0011 ' ГЭСН01-02-001-01
                End If

            ' --- ОБЕЗПЫЛИВАНИЕ / ОКРАСКА / ИЗОЛЯЦИЯ ---
            Case InStr(1, rName, "Обеспыливание", vbTextCompare) > 0 Or InStr(1, rName, "Окраска", vbTextCompare) > 0 Or InStr(1, rName, "Ораска", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м2"
                ws.Cells(i, 5).Value = 0.045 ' ГЭСН13-03-002-04
            Case InStr(1, rName, "Изоляция битумно-уретановым", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м"
                ws.Cells(i, 5).Value = 0.38 ' ГЭСН22-01-011-03

            ' --- БУРЕНИЕ И КРЕПЛЕНИЕ СКВАЖИН ---
            Case InStr(1, rName, "бурение скважин", vbTextCompare) > 0 Or InStr(1, rName, "бурения до", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м"
                ws.Cells(i, 5).Value = 0.85 ' ГЭСН04-01-001-01
            Case InStr(1, rName, "Крепление скважины", vbTextCompare) > 0 Or InStr(1, rName, "Свободный спуск", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м"
                ws.Cells(i, 5).Value = 0.42 ' ГЭСН04-01-030-01

            ' --- ДЕМОНТАЖНЫЕ И ПРОЧИЕ СТРОИТЕЛЬНЫЕ ПОЗИЦИИ ---
            Case InStr(1, rName, "Демонтаж стальных", vbTextCompare) > 0 Or InStr(1, rName, "демонтаж внутриплощадочных", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "км"
                ws.Cells(i, 5).Value = 184.5 ' ГЭСН22-05-001-01
            Case InStr(1, rName, "Демонтаж кабельной", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м"
                ws.Cells(i, 5).Value = 0.024 ' ГЭСНм08-02-141-01
            Case InStr(1, rName, "Погрузка излишек", vbTextCompare) > 0 Or InStr(1, rName, "Погрузка лишнего", vbTextCompare) > 0 Or InStr(1, rName, "Погрузка грунта", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                ws.Cells(i, 5).Value = 0.0115 ' ГЭСН01-01-012-01
            Case InStr(1, rName, "Разборка покрытий", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                ws.Cells(i, 5).Value = 1.14 ' ГЭСН27-03-004-01
            Case InStr(1, rName, "Сбор строительного мусора", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "м3"
                ws.Cells(i, 5).Value = 0.045 ' ГЭСН01-01-033-01
            Case InStr(1, rName, "Разводка по устройствам", vbTextCompare) > 0
                ws.Cells(i, 4).Value = "100 жил"
                ws.Cells(i, 5).Value = 14.8 ' ГЭСНм08-03-574-01

            ' --- ДЕФОЛТНОЕ ЗНАЧЕНИЕ ДЛЯ ОСТАЛЬНЫХ ПОЗИЦИЙ ---
            Case Else
                If ws.Cells(i, 4).Value = "" Then ws.Cells(i, 4).Value = "шт"
                ws.Cells(i, 5).Value = 1.25
        End Select
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Данные успешно внесены.", vbInformation, "Профессиональный режим"
End Sub
