Attribute VB_Name = "Module2"
Sub SumFilteredData()
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalSum As Double
    
    ' Указываем листы
    Set wsData = ThisWorkbook.Worksheets("Задание 1") ' Лист с данными
    Set wsResult = ThisWorkbook.Worksheets("Задание 1.1") ' Лист для результатов
    
    ' Находим последнюю строку в столбце AD
    lastRow = wsData.Cells(wsData.Rows.Count, "AD").End(xlUp).Row
    
    ' Инициализируем сумму
    totalSum = 0
    
    ' Проходим по всем строкам данных
    For i = 2 To lastRow ' Предполагаем заголовки в первой строке
        ' Проверяем условия фильтрации
        If wsData.Cells(i, "V").Value = "Смена. Доп" And _
           wsData.Cells(i, "AB").Value = "b2c СГ Проблемы с доставкой" Then
            
            ' Суммируем только числовые значения
            If IsNumeric(wsData.Cells(i, "AD").Value) Then
                totalSum = totalSum + wsData.Cells(i, "AD").Value
            End If
        End If
    Next i
    
    ' Записываем результат
    wsResult.Range("G1").Value = totalSum
    MsgBox "Сумма: " & totalSum & vbCrLf & "Результат записан в A1 листа 'Результат'"
End Sub

