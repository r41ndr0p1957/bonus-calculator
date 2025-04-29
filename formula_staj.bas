Attribute VB_Name = "Module3"
Sub CalculateAllRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    For i = 2 To lastRow
        ' Для первой формулы результат выводим в столбец I
        If IsDate(ws.Cells(i, 8).Value) Then
            ws.Cells(i, 9).Value = Round((Date - ws.Cells(i, 8).Value) / 30.5, 2)
        End If
        
        ' Для первой формулы результат выводим в столбец J
        If IsDate(ws.Cells(i, 8).Value) And IsDate(ws.Range("F2").Value) Then
            Dim days As Long
            days = DateDiff("d", ws.Cells(i, 8).Value, ws.Range("F2").Value)
            ws.Cells(i, 10).Value = Round(days / 30.5, 2)
        End If
    Next i
End Sub
