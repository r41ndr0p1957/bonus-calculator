Attribute VB_Name = "Module1"
Sub CalculateShiftHours()
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, key As Variant '
    Dim login As String, shiftType As String, shiftDate As Date
    Dim startDateTime As Date, endDateTime As Date
    Dim totalHours As Double
    
    ' Создаем коллекцию для хранения данных (логин + дата -> сумма часов)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Определение рабочих листов
    Set wsSource = ThisWorkbook.Sheets("Задание 1") ' Тут название исходного листа
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Результат").Delete
    Application.DisplayAlerts = True
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "Задание 1.1"
    
    ' Находим последнюю заполненную строку в исходном листе
    lastRow = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row
    
    ' Проходим по всем строкам, пропускаем строку 1 (заголовок), собираем данные
    For i = 2 To lastRow
        login = wsSource.Cells(i, "G").Value
        shiftType = wsSource.Cells(i, "V").Value
        shiftDate = wsSource.Cells(i, "W").Value
        
        ' Проверка типа смены
        If shiftType = "Смена. Основная" Or shiftType = "Смена. Доп" Or _
           shiftType = "Смена. Отработка" Or shiftType = "Сегмент смены" Then
            
            ' Старт + конец
            startDateTime = wsSource.Cells(i, "W").Value + wsSource.Cells(i, "X").Value
            endDateTime = wsSource.Cells(i, "Y").Value + wsSource.Cells(i, "Z").Value
            
            ' Считаем часики
            totalHours = (endDateTime - startDateTime) * 24
            
            ' Ключ (логин + дата)
            key = login & "|" & shiftDate
            
            ' Добавляем в коллекцию
            If dict.Exists(key) Then
                dict(key) = dict(key) + totalHours
            Else
                dict.Add key, totalHours
            End If
        End If
    Next i
    
    ' Выводим результат в новый лист
    With wsResult
        ' Заголовки столбцов
        .Cells(1, 1).Value = "Логин"
        .Cells(1, 2).Value = "Дата"
        .Cells(1, 3).Value = "Сумма часов"
        
        ' Заполняем данные
        Dim rowIndex As Long
        rowIndex = 2
        For Each key In dict.Keys ' Теперь key объявлен как Variant
            login = Split(key, "|")(0)
            shiftDate = Split(key, "|")(1)
            .Cells(rowIndex, 1).Value = login
            .Cells(rowIndex, 2).Value = shiftDate
            .Cells(rowIndex, 3).Value = dict(key)
            rowIndex = rowIndex + 1
        Next key
        
        ' Форматируем дату и числа
        .Columns("B:B").NumberFormat = "dd.mm.yyyy"
        .Columns("C:C").NumberFormat = "0.00"
        
        ' Автоподбор высоты столбцов
        .Columns("A:C").AutoFit
    End With
    
    MsgBox "Обработка завершена! Результат в листе 'Результат'."
End Sub

