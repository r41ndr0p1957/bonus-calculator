Attribute VB_Name = "Module1"
Sub CalculatePremiums()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim maxC As Double, totalOperators As Long
    Dim dataArr() As Variant, resultArr() As Variant
    Dim rankCounts(1 To 6) As Long, currentRank As Integer, cumulative As Long
    Dim totalAssigned As Long, overflow As Long

    Set wsSource = ThisWorkbook.Worksheets("Премия")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Нет данных для обработки", vbExclamation
        Exit Sub
    End If

    maxC = Application.WorksheetFunction.Max(wsSource.Range("C2:C" & lastRow))
    If maxC = 0 Then maxC = 1

    dataArr = wsSource.Range("A2:E" & lastRow).Value
    totalOperators = UBound(dataArr, 1)
    ReDim resultArr(1 To totalOperators, 1 To 6)

    For i = 1 To totalOperators
        resultArr(i, 3) = Round(dataArr(i, 3) / maxC * 0.1, 4)
        resultArr(i, 4) = Round(dataArr(i, 4) / 5 * 0.4, 4)
        resultArr(i, 5) = Round(dataArr(i, 5) / 100 * 0.5, 4)
        resultArr(i, 6) = Round(resultArr(i, 3) + resultArr(i, 4) + resultArr(i, 5), 4)
        resultArr(i, 2) = dataArr(i, 1)
    Next i

    Application.ScreenUpdating = False
    Set wsDest = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsDest.Name = "TempSort"
    wsDest.Range("A1:F1").Value = Array("Ранг", "Логин", "Вес C", "Вес D", "Вес E", "Итог")
    wsDest.Range("A2").Resize(totalOperators, 6).Value = resultArr

    With wsDest.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDest.Range("F2:F" & totalOperators + 1), Order:=xlDescending
        .SetRange wsDest.Range("A2:F" & totalOperators + 1)
        .Header = xlNo
        .Apply
    End With

    resultArr = wsDest.Range("A2:F" & totalOperators + 1).Value
    Application.DisplayAlerts = False
    wsDest.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    rankCounts(1) = Application.RoundUp(totalOperators * 0.05, 0)
    rankCounts(2) = Application.RoundUp(totalOperators * 0.1, 0)
    rankCounts(3) = Application.RoundUp(totalOperators * 0.15, 0)
    rankCounts(4) = Application.RoundUp(totalOperators * 0.2, 0)
    rankCounts(5) = Application.RoundUp(totalOperators * 0.25, 0)
    totalAssigned = Application.Sum(rankCounts)

    If totalAssigned > totalOperators Then
        overflow = totalAssigned - totalOperators
        rankCounts(5) = rankCounts(5) - overflow
        If rankCounts(5) < 0 Then
            rankCounts(4) = rankCounts(4) + rankCounts(5)
            rankCounts(5) = 0
        End If
        totalAssigned = Application.Sum(rankCounts)
    End If

    rankCounts(6) = totalOperators - totalAssigned
    If rankCounts(6) < 0 Then rankCounts(6) = 0


    currentRank = 1
    cumulative = 1
    For i = 1 To 6
        For j = 1 To rankCounts(i)
            If cumulative > totalOperators Then Exit For
            resultArr(cumulative, 1) = currentRank
            cumulative = cumulative + 1
        Next j
        currentRank = currentRank + 1
    Next i

    Set wsDest = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsDest.Name = "Результаты"
    wsDest.Range("A1:F1").Value = Array("Ранг", "Логин", "Вес сделки", "Вес CSAT", "Вес QQ", "Итоговый балл")
    wsDest.Range("A2").Resize(totalOperators, 6).Value = resultArr
    wsDest.Columns.AutoFit

    MsgBox "Готово! Результаты на листе '" & wsDest.Name & "'", vbInformation
End Sub
