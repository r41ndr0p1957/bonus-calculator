Attribute VB_Name = "Module2"
Sub SumFilteredData()
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalSum As Double
    
    ' ��������� �����
    Set wsData = ThisWorkbook.Worksheets("������� 1") ' ���� � �������
    Set wsResult = ThisWorkbook.Worksheets("������� 1.1") ' ���� ��� �����������
    
    ' ������� ��������� ������ � ������� AD
    lastRow = wsData.Cells(wsData.Rows.Count, "AD").End(xlUp).Row
    
    ' �������������� �����
    totalSum = 0
    
    ' �������� �� ���� ������� ������
    For i = 2 To lastRow ' ������������ ��������� � ������ ������
        ' ��������� ������� ����������
        If wsData.Cells(i, "V").Value = "�����. ���" And _
           wsData.Cells(i, "AB").Value = "b2c �� �������� � ���������" Then
            
            ' ��������� ������ �������� ��������
            If IsNumeric(wsData.Cells(i, "AD").Value) Then
                totalSum = totalSum + wsData.Cells(i, "AD").Value
            End If
        End If
    Next i
    
    ' ���������� ���������
    wsResult.Range("G1").Value = totalSum
    MsgBox "�����: " & totalSum & vbCrLf & "��������� ������� � A1 ����� '���������'"
End Sub

