Attribute VB_Name = "Mod_ImportManageDivide"
Option Explicit

Sub DivideManage()

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    Dim insDate As String
    
    Application.ScreenUpdating = False
    
    ClearManageDivide
    
    With shtManageData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
        For i = 2 To endRow
            man.ID = .Cells(i, 1)
            man.�������� = Trim(.Cells(i, 2))
            man.�з�1 = Trim(.Cells(i, 3))
            man.�з�2 = Trim(.Cells(i, 4))
            man.������ȣ = .Cells(i, 5)
            man.�ŷ�ó = .Cells(i, 6)
            man.ǰ�� = .Cells(i, 7)
            man.���� = .Cells(i, 8)
            man.�԰� = .Cells(i, 9)
            man.�ܰ� = .Cells(i, 10)
            man.�ݾ� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            man.�߷� = .Cells(i, 13)
            man.���� = .Cells(i, 14)
            man.���� = .Cells(i, 15)
            man.���� = .Cells(i, 16)
            man.���� = .Cells(i, 17)
            man.�԰� = .Cells(i, 18)
            man.��ǰ = .Cells(i, 19)
            man.���� = .Cells(i, 20)
            man.��꼭 = .Cells(i, 21)
            man.���� = .Cells(i, 22)
            man.����� = .Cells(i, 23)
            man.�ΰ��� = .Cells(i, 24)
            man.������� = .Cells(i, 25)
            
            '������� ����
'            insDate = ""
'            If man.�з�2 = "����" Or (man.�������� = "����" And Len(man.������ȣ) >= 10) Then
'                If man.�з�2 = "����" Then
'                    insDate = man.����
'                Else
'                    If man.���� <> "" Then
'                        insDate = man.����
'                    ElseIf man.���� <> "" Then
'                        insDate = man.����
'                    ElseIf man.��꼭 <> "" Then
'                        insDate = man.��꼭
'                    ElseIf man.���� <> "" Then
'                        insDate = man.����
'                    End If
'                End If
'
'                If insDate = "" Then
'                    pos = InStr(man.������ȣ, "-")
'                    If pos > 0 Then
'                        M = Mid(man.������ȣ, pos - 4, 2)
'                        D = Mid(man.������ȣ, pos - 2, 2)
'                        If IsNumeric(M) And IsNumeric(D) Then
'                            Y = year(man.�������)
'                            insDate = DateSerial(Y, M, D)
'                        End If
'                    End If
'                End If
'
'                If insDate <> "" Then
'                    man.������� = insDate
'                End If
'            End If
               
            If man.�з�2 = "����" Then
                '�����̸� ���� ���̺� ���
                Insert_Record shtAcceptedData, man.ID, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, man.����, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������
                '���ֹ��� ���̺��� ���
                Insert_Record shtOrderData, man.ID, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, man.����, man.�԰�, man.����, man.����, man.�ܰ�, man.�ݾ�, man.�߷�, _
                              man.����, man.����, man.����, man.�԰�, man.��ǰ, man.����, man.��꼭, man.����, man.�����, , man.�ΰ���, man.�������

            ElseIf man.�������� = "����" And Len(man.������ȣ) >= 10 Then
                '�����̸鼭 ������ȣ�� ������ ���ֹ��� ���̺� ���
                '������ ��쿡�� �з�1�� �������� �ʵ忡 ����
                Insert_Record shtOrderData, man.ID, , man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, man.����, man.�԰�, man.����, man.����, man.�ܰ�, man.�ݾ�, man.�߷�, _
                              man.����, man.����, man.����, man.�԰�, man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�з�1, man.�ΰ���, man.�������
            Else
                '�� �� ���� ���� ��� ���̺� ���
                Insert_Record shtOperatingData, man.ID, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, man.�ݾ�, man.����, man.��꼭, man.����, man.�ΰ���, man.�������
            End If
            
        Next
    End With
End Sub

Sub ClearManageDivide()

    Dim endCol, endRow As Long
    
    With shtAcceptedData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtOrderData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtOperatingData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub
