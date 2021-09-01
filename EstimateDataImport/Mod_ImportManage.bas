Attribute VB_Name = "Mod_ImportManage"
Option Explicit

'���� ����ü
Type Manage
    ID As Long
    �������� As String
    �з�1 As String
    �з�2 As String
    ������ȣ As String
    �ŷ�ó As String
    ǰ�� As String
    ���� As String
    �԰� As String
    �ܰ� As String
    �ݾ� As String
    ���� As String
    �߷� As String
    ���� As String
    ���� As String
    ���� As String
    ���� As String
    �԰� As String
    ��ǰ As String
    ���� As String
    ��꼭 As String
    ���� As String
    ����� As String
    �ΰ��� As String
    ������� As String
    �������� As String
End Type

Sub ImportManage()

    Dim WB As Workbook
    Dim WS As Worksheet:
    Dim i As Long
    Dim j As Long
    Dim strWS As String
    Dim manageFileList(1) As Variant
    Dim importCount As Long
    Dim pos As Long
    Dim M As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ClearManageData
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2005����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            pos = InStr(WS.Name, "��")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 5 And M <= 6 Then
                        importCount = ImportManageData_Type9(WS, WS.Name, 2005)
                    Else
                        importCount = ImportManageData_Type8(WS, WS.Name, 2005)
                    End If
                End If
            End If
        End If
    Next
    WB.Close

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2006����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            pos = InStr(WS.Name, "��")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 2 Then
                        importCount = ImportManageData_Type7(WS, WS.Name, 2006)
                    Else
                        importCount = ImportManageData_Type6(WS, WS.Name, 2006)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2007����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            pos = InStr(WS.Name, "��")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 3 Then
                        importCount = ImportManageData_Type6(WS, WS.Name, 2007)
                    Else
                        importCount = ImportManageData_Type5(WS, WS.Name, 2007)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2008����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2008)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2009����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2009)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2010����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2010)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2011����.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            pos = InStr(WS.Name, "��")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 2 Then
                        importCount = ImportManageData_Type4(WS, WS.Name, 2011)
                    Else
                        importCount = ImportManageData_Type3(WS, WS.Name, 2011)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2012����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type3(WS, WS.Name, 2012)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\2013����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type2(WS, WS.Name, 2013)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2013����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            pos = InStr(WS.Name, "��")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If M >= 1 And M <= 10 Then
                    importCount = ImportManageData_Type2(WS, WS.Name, 2013)
                Else
                    importCount = ImportManageData_Type1(WS, WS.Name, 2013)
                End If
            End If
            
        End If
    Next
    WB.Close


    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2014����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2014)
        End If
    Next
    WB.Close

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2015����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2015)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2016����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2016)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2017����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2017)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2018����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2018)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2019����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2019)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2020����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2020)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-��������\����-2021����.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2021)
        End If
    Next
    WB.Close

End Sub

Sub ClearManageData()
    Dim endCol, endRow As Long
    
    With shtManageData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtManageMemoData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub


Function ImportManageData_Type1(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    With WS
        endCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            man.�з�1 = .Cells(i, 2)
            man.�з�2 = .Cells(i, 3)
            man.������ȣ = .Cells(i, 4)
            man.�ŷ�ó = .Cells(i, 5)
            man.ǰ�� = .Cells(i, 6)
            man.���� = .Cells(i, 7)
            man.�԰� = .Cells(i, 8)
            man.�ܰ� = .Cells(i, 9)
            man.�ݾ� = .Cells(i, 10)
            man.���� = .Cells(i, 11)
            man.�߷� = .Cells(i, 12)
            man.���� = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.���� = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.���� = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.���� = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.���� = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.���� = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.���� = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.���� = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.���� = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.��꼭 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.���� = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 30)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            man.�ΰ��� = .Cells(i, 31)
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type1 = importCount
End Function

Function ImportManageData_Type2(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            man.�з�1 = .Cells(i, 2)
            man.�з�2 = .Cells(i, 3)
            man.������ȣ = .Cells(i, 4)
            man.�ŷ�ó = .Cells(i, 5)
            man.ǰ�� = .Cells(i, 6)
            man.���� = .Cells(i, 7)
            man.�԰� = .Cells(i, 8)
            man.�ܰ� = .Cells(i, 9)
            man.�ݾ� = .Cells(i, 10)
            man.���� = .Cells(i, 11)
            man.�߷� = .Cells(i, 12)
            man.���� = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.���� = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.���� = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.���� = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.���� = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.���� = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.���� = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.���� = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.���� = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.��꼭 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.���� = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 30)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            man.�ΰ��� = .Cells(i, 31)
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type2 = importCount
End Function

Function ImportManageData_Type3(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        Else
            regDate = DateSerial(Y, M, 1)
        End If
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            man.�з�1 = .Cells(i, 2)
            man.�з�2 = .Cells(i, 3)
            man.������ȣ = .Cells(i, 4)
            man.�ŷ�ó = .Cells(i, 5)
            man.ǰ�� = .Cells(i, 6)
            man.���� = .Cells(i, 7)
            man.�԰� = .Cells(i, 8)
            man.�ܰ� = .Cells(i, 9)
            man.�ݾ� = .Cells(i, 10)
            man.���� = .Cells(i, 11)
            man.�߷� = .Cells(i, 12)
            man.���� = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.���� = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.���� = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.���� = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.���� = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.���� = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.���� = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.���� = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.���� = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.��꼭 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.���� = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 30)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
                
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type3 = importCount
End Function

Function ImportManageData_Type4(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ���� - ������ ��Ʈ�� ���� �ϳ� �߰��ؼ� �ذ���
            man.�з�1 = .Cells(i, 2)
            man.�з�2 = .Cells(i, 3)
            man.������ȣ = .Cells(i, 4)
            man.�ŷ�ó = .Cells(i, 5)
            man.ǰ�� = .Cells(i, 6)
            man.���� = .Cells(i, 7)
            man.�԰� = .Cells(i, 8)
            man.�ܰ� = .Cells(i, 9)
            man.�ݾ� = .Cells(i, 10)
            man.���� = .Cells(i, 11)
            man.�߷� = .Cells(i, 12)
            man.���� = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.���� = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.���� = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.���� = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.���� = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.���� = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.���� = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.���� = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.���� = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.��꼭 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.���� = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 30)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type4 = importCount
End Function

Function ImportManageData_Type5(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ����
            man.�з�1 = ""
            man.�з�2 = .Cells(i, 2)
            man.������ȣ = .Cells(i, 3)
            man.�ŷ�ó = .Cells(i, 4)
            man.ǰ�� = .Cells(i, 5)
            man.���� = .Cells(i, 6)
            man.�԰� = .Cells(i, 7)
            man.�ܰ� = .Cells(i, 8)
            man.�ݾ� = .Cells(i, 9)
            man.���� = .Cells(i, 10)
            man.�߷� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.���� = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.���� = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.���� = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.���� = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.���� = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.���� = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.���� = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.���� = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.��꼭 = ""
            End If
                
            If .Cells(i, 27) <> "" And .Cells(i, 28) <> "" And IsNumeric(.Cells(i, 27)) And IsNumeric(.Cells(i, 28)) Then
                man.���� = DateSerial(Y, .Cells(i, 27), .Cells(i, 28))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 29)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type5 = importCount
End Function

Function ImportManageData_Type6(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ����
            man.�з�1 = ""
            man.�з�2 = .Cells(i, 2)
            man.������ȣ = .Cells(i, 3)
            man.�ŷ�ó = .Cells(i, 4)
            man.ǰ�� = .Cells(i, 5)
            man.���� = .Cells(i, 6)
            man.�԰� = .Cells(i, 7)
            man.�ܰ� = .Cells(i, 8)
            man.�ݾ� = .Cells(i, 9)
            man.���� = .Cells(i, 10)
            man.�߷� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.���� = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.���� = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.���� = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.���� = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.���� = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.���� = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.��꼭 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.���� = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.���� = ""
            End If
            
            man.����� = .Cells(i, 27)
            If man.����� <> "" Then
                If IsNumeric(man.�����) Then
                    M = CLng(man.�����)
                    If M >= 1 And M <= 12 Then
                        man.����� = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type6 = importCount
End Function

Function ImportManageData_Type7(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2�� �Ӹ��� 9��, ������ 11������ �м�
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ����
            man.�з�1 = ""
            man.�з�2 = .Cells(i, 2)
            man.������ȣ = .Cells(i, 3)
            man.�ŷ�ó = .Cells(i, 4)
            man.ǰ�� = .Cells(i, 5)
            man.���� = .Cells(i, 6)
            man.�԰� = .Cells(i, 7)
            man.�ܰ� = .Cells(i, 8)
            man.�ݾ� = .Cells(i, 9)
            man.���� = .Cells(i, 10)
            man.�߷� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.���� = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.���� = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.���� = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.���� = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.���� = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.���� = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.��꼭 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.���� = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.���� = ""
            End If
            
            'Type7�� ����� ����
            man.����� = ""
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type7 = importCount
End Function

Function ImportManageData_Type8(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type8�� �Ӹ��� 8��, ������ 10������ �м�
    With WS
        endCol = .Cells(8, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 10 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ����
            man.�з�1 = ""
            man.�з�2 = .Cells(i, 2)
            man.������ȣ = .Cells(i, 3)
            man.�ŷ�ó = .Cells(i, 4)
            man.ǰ�� = .Cells(i, 5)
            man.���� = .Cells(i, 6)
            man.�԰� = .Cells(i, 7)
            man.�ܰ� = .Cells(i, 8)
            man.�ݾ� = .Cells(i, 9)
            man.���� = .Cells(i, 10)
            man.�߷� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.���� = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.���� = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.���� = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.���� = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.���� = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.���� = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.��꼭 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.���� = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.���� = ""
            End If
            
            'Type7�� ����� ����
            man.����� = ""
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
        
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type8 = importCount
End Function

Function ImportManageData_Type9(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type9�� �Ӹ��� 3��, ������ 5������ �м�
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 5 To endRow
 
        With WS
            man.�������� = .Cells(i, 1)
            'Type4�� �з�1 ����
            man.�з�1 = ""
            man.�з�2 = .Cells(i, 2)
            man.������ȣ = .Cells(i, 3)
            man.�ŷ�ó = .Cells(i, 4)
            man.ǰ�� = .Cells(i, 5)
            man.���� = .Cells(i, 6)
            man.�԰� = .Cells(i, 7)
            man.�ܰ� = .Cells(i, 8)
            man.�ݾ� = .Cells(i, 9)
            man.���� = .Cells(i, 10)
            man.�߷� = .Cells(i, 11)
            man.���� = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.���� = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.���� = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.���� = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.���� = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.�԰� = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.�԰� = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.��ǰ = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.��ǰ = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.���� = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.���� = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.��꼭 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.��꼭 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.���� = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.���� = ""
            End If
            
            'Type7�� ����� ����
            man.����� = ""
            
            'Type3�� �ΰ��� ����
            man.�ΰ��� = ""
            man.������� = regDate
        End With
        
        '���� ���̺� ���
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.��������, man.�з�1, man.�з�2, man.������ȣ, man.�ŷ�ó, man.ǰ��, _
                      man.����, man.�԰�, man.�ܰ�, man.�ݾ�, man.����, man.�߷�, man.����, man.����, man.����, man.����, man.�԰�, _
                      man.��ǰ, man.����, man.��꼭, man.����, man.�����, man.�ΰ���, man.�������, man.��������
        
        importCount = importCount + 1
        
       
        '�޸� ���̺� ���
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.������ȣ, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type9 = importCount
End Function


