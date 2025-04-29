Attribute VB_Name = "Module1"
Sub CalculateShiftHours()
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, key As Variant '
    Dim login As String, shiftType As String, shiftDate As Date
    Dim startDateTime As Date, endDateTime As Date
    Dim totalHours As Double
    
    ' ������� ��������� ��� �������� ������ (����� + ���� -> ����� �����)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ����������� ������� ������
    Set wsSource = ThisWorkbook.Sheets("������� 1") ' ��� �������� ��������� �����
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("���������").Delete
    Application.DisplayAlerts = True
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "������� 1.1"
    
    ' ������� ��������� ����������� ������ � �������� �����
    lastRow = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row
    
    ' �������� �� ���� �������, ���������� ������ 1 (���������), �������� ������
    For i = 2 To lastRow
        login = wsSource.Cells(i, "G").Value
        shiftType = wsSource.Cells(i, "V").Value
        shiftDate = wsSource.Cells(i, "W").Value
        
        ' �������� ���� �����
        If shiftType = "�����. ��������" Or shiftType = "�����. ���" Or _
           shiftType = "�����. ���������" Or shiftType = "������� �����" Then
            
            ' ����� + �����
            startDateTime = wsSource.Cells(i, "W").Value + wsSource.Cells(i, "X").Value
            endDateTime = wsSource.Cells(i, "Y").Value + wsSource.Cells(i, "Z").Value
            
            ' ������� ������
            totalHours = (endDateTime - startDateTime) * 24
            
            ' ���� (����� + ����)
            key = login & "|" & shiftDate
            
            ' ��������� � ���������
            If dict.Exists(key) Then
                dict(key) = dict(key) + totalHours
            Else
                dict.Add key, totalHours
            End If
        End If
    Next i
    
    ' ������� ��������� � ����� ����
    With wsResult
        ' ��������� ��������
        .Cells(1, 1).Value = "�����"
        .Cells(1, 2).Value = "����"
        .Cells(1, 3).Value = "����� �����"
        
        ' ��������� ������
        Dim rowIndex As Long
        rowIndex = 2
        For Each key In dict.Keys ' ������ key �������� ��� Variant
            login = Split(key, "|")(0)
            shiftDate = Split(key, "|")(1)
            .Cells(rowIndex, 1).Value = login
            .Cells(rowIndex, 2).Value = shiftDate
            .Cells(rowIndex, 3).Value = dict(key)
            rowIndex = rowIndex + 1
        Next key
        
        ' ����������� ���� � �����
        .Columns("B:B").NumberFormat = "dd.mm.yyyy"
        .Columns("C:C").NumberFormat = "0.00"
        
        ' ���������� ������ ��������
        .Columns("A:C").AutoFit
    End With
    
    MsgBox "��������� ���������! ��������� � ����� '���������'."
End Sub

