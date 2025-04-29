
Sub CalculateShiftHours()
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, key As Variant '
    Dim login As String, shiftType As String, shiftDate As Date
    Dim startDateTime As Date, endDateTime As Date
    Dim totalHours As Double
    
    ' Ñîçäàåì êîëëåêöèþ äëÿ õðàíåíèÿ äàííûõ (ëîãèí + äàòà -> ñóììà ÷àñîâ)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Îïðåäåëåíèå ðàáî÷èõ ëèñòîâ
    Set wsSource = ThisWorkbook.Sheets("Çàäàíèå 1") ' Òóò íàçâàíèå èñõîäíîãî ëèñòà
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Ðåçóëüòàò").Delete
    Application.DisplayAlerts = True
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "Çàäàíèå 1.1"
    
    ' Íàõîäèì ïîñëåäíþþ çàïîëíåííóþ ñòðîêó â èñõîäíîì ëèñòå
    lastRow = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row
    
    ' Ïðîõîäèì ïî âñåì ñòðîêàì, ïðîïóñêàåì ñòðîêó 1 (çàãîëîâîê), ñîáèðàåì äàííûå
    For i = 2 To lastRow
        login = wsSource.Cells(i, "G").Value
        shiftType = wsSource.Cells(i, "V").Value
        shiftDate = wsSource.Cells(i, "W").Value
        
        ' Ïðîâåðêà òèïà ñìåíû
        If shiftType = "Ñìåíà. Îñíîâíàÿ" Or shiftType = "Ñìåíà. Äîï" Or _
           shiftType = "Ñìåíà. Îòðàáîòêà" Or shiftType = "Ñåãìåíò ñìåíû" Then
            
            ' Ñòàðò + êîíåö
            startDateTime = wsSource.Cells(i, "W").Value + wsSource.Cells(i, "X").Value
            endDateTime = wsSource.Cells(i, "Y").Value + wsSource.Cells(i, "Z").Value
            
            ' Ñ÷èòàåì ÷àñèêè
            totalHours = (endDateTime - startDateTime) * 24
            
            ' Êëþ÷ (ëîãèí + äàòà)
            key = login & "|" & shiftDate
            
            ' Äîáàâëÿåì â êîëëåêöèþ
            If dict.Exists(key) Then
                dict(key) = dict(key) + totalHours
            Else
                dict.Add key, totalHours
            End If
        End If
    Next i
    
    ' Âûâîäèì ðåçóëüòàò â íîâûé ëèñò
    With wsResult
        ' Çàãîëîâêè ñòîëáöîâ
        .Cells(1, 1).Value = "Ëîãèí"
        .Cells(1, 2).Value = "Äàòà"
        .Cells(1, 3).Value = "Ñóììà ÷àñîâ"
        
        ' Çàïîëíÿåì äàííûå
        Dim rowIndex As Long
        rowIndex = 2
        For Each key In dict.Keys ' Òåïåðü key îáúÿâëåí êàê Variant
            login = Split(key, "|")(0)
            shiftDate = Split(key, "|")(1)
            .Cells(rowIndex, 1).Value = login
            .Cells(rowIndex, 2).Value = shiftDate
            .Cells(rowIndex, 3).Value = dict(key)
            rowIndex = rowIndex + 1
        Next key
        
        ' Ôîðìàòèðóåì äàòó è ÷èñëà
        .Columns("B:B").NumberFormat = "dd.mm.yyyy"
        .Columns("C:C").NumberFormat = "0.00"
        
        ' Àâòîïîäáîð âûñîòû ñòîëáöîâ
        .Columns("A:C").AutoFit
    End With
    
    MsgBox "Îáðàáîòêà çàâåðøåíà! Ðåçóëüòàò â ëèñòå 'Ðåçóëüòàò'."
End Sub

