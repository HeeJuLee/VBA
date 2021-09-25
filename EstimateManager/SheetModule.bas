Attribute VB_Name = "SheetModule"
Option Explicit

Sub UpdateShtOrderField(orderId, fieldName, fieldValue)
    Dim findRow, colNo As Long
    
    findRow = isExistInSheet(shtOrderAdmin.Range("B6"), orderId)
    If findRow > 0 Then
        Select Case fieldName
            Case "�з�1"
                colNo = 7
            Case "�ŷ�ó"
                colNo = 8
            Case "ǰ��"
                colNo = 9
            Case "����"
                colNo = 10
            Case "�԰�"
                colNo = 11
            Case "����"
                colNo = 12
            Case "����"
                colNo = 13
            Case "�ܰ�"
                colNo = 14
            Case "�ݾ�"
                colNo = 15
            Case "�߷�"
                colNo = 16
            Case "����"
                colNo = 17
            Case "����"
                colNo = 18
            Case "����"
                colNo = 19
            Case "�԰�"
                colNo = 20
            Case "��ǰ"
                colNo = 21
            Case "����"
                colNo = 22
            Case "��꼭"
                colNo = 23
            Case "����"
                colNo = 24
            Case "������"
                colNo = 25
            Case "��������"
                colNo = 26
            Case "�ΰ���"
                colNo = 27
            Case "�������"
                colNo = 28
            Case "��������"
                colNo = 29
        End Select
        
        If colNo > 0 Then
            shtOrderAdmin.Cells(findRow, colNo).value = fieldValue
        End If
    End If
End Sub


