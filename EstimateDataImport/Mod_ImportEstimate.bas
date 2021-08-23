Attribute VB_Name = "Mod_ImportEstimate"
Option Explicit

'���� �����׸�
Type Production
    ID As String
    ID_���� As String
    �׸� As String
    ��� As String
    �޸� As String
End Type

'���� ����ü
Type Estimate
    ID As Long
    ������ȣ As String
    �����ȣ As String
    �ŷ�ó As String
    ����� As String
    ǰ�� As String
    �԰� As String
    ���� As String
    ���� As String
    �����ܰ� As String
    �����ݾ� As String
    ������ As String
    ������ As String
    ������ As String
    ��ǰ�� As String
    ������ As String
    ����� As String
    �̸� As String
    ���� As String
    �ΰǺ� As String
    ��Ÿ As String
    ���డ As String
    �����ݾ� As String
    ���� As String
    ������ As String
    ���ֱݾ� As String
    �������� As String
    ������� As String
    �������� As String
End Type

Sub ImportEstimate()

    Dim WB As Workbook
    Dim WS As Worksheet:
    Dim i As Long
    Dim j As Long
    Dim strWS As String
    Dim estimateFileList(1) As Variant
    Dim importCount As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ClearEstimateData
    
'    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2013.xlsx")
'    For Each WS In WB.Worksheets
'        If WS.Name Like "**��" Then
'            importCount = ImportEstimateData_Type1(WS, WS.Name, 2013)
'        End If
'    Next
'    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2014.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2014)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2015.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2015)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2016.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2016)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2017.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2017)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2018.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2018)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2019.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2019)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2020.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2020)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\����-������������\����-��������2021.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**��" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2021)
        End If
    Next
    WB.Close

End Sub

Sub ClearEstimateData()

    Dim endCol, endRow As Long
    
    With shtEstimateData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtProductionData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtEstimateMemoData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub

Function ImportEstimateData_Type1(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As Estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "��")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1��~12���� �ƴϸ� ����
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type1 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.������ȣ = .Cells(i, 1)
            est.�����ȣ = .Cells(i, 2)
            est.�ŷ�ó = .Cells(i, 3)
            est.����� = .Cells(i, 4)
            est.ǰ�� = .Cells(i, 5)
            
            '�������� ���� ���� �����ִ� �����̹Ƿ� ����
            If est.ǰ�� = "" Then
                GoTo NextIteration
            End If
            
            est.�԰� = .Cells(i, 6)
            est.���� = .Cells(i, 7)
            est.���� = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! ���� ����
                est.�����ܰ� = .Cells(i, 9)
            Else
                est.�����ܰ� = ""
            End If
            est.�����ݾ� = .Cells(i, 10)
            
            est.������ = Get_Date_Convert(.Cells(i, 11), year)
            est.������ = Get_Date_Convert(.Cells(i, 12), year)
            est.������ = Get_Date_Convert(.Cells(i, 13), year)
            est.��ǰ�� = Get_Date_Convert(.Cells(i, 14), year)
            est.������ = Get_Date_Convert(.Cells(i, 15), year)
            
'            est.������ = .Cells(i, 11)
'            If Len(est.������) = 5 Then
'                est.������ = DateSerial(Y, Left(est.������, 2), Right(est.������, 2))
'            End If
'            est.������ = .Cells(i, 12)
'            If Len(est.������) = 5 Then
'                est.������ = DateSerial(Y, Left(est.������, 2), Right(est.������, 2))
'            End If
'            est.������ = .Cells(i, 13)
'            If Len(est.������) = 5 Then
'                est.������ = DateSerial(Y, Left(est.������, 2), Right(est.������, 2))
'            End If
'            est.��ǰ�� = .Cells(i, 14)
'            If Len(est.��ǰ��) = 5 Then
'                est.��ǰ�� = DateSerial(Y, Left(est.��ǰ��, 2), Right(est.��ǰ��, 2))
'            End If
'            est.������ = .Cells(i, 15)
'            If Len(est.������) = 5 Then
'                est.������ = DateSerial(Y, Left(est.������, 2), Right(est.������, 2))
'            End If
            
            '��������׸� ���
            If .Cells(i, 16) <> "" Then
                prod.�׸� = "�����"
                prod.��� = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.�޸� = "" Else prod.�޸� = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.������ȣ, prod.�׸�, prod.���, prod.�޸�, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.�׸� = "�̸�"
                prod.��� = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.�޸� = "" Else prod.�޸� = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.������ȣ, prod.�׸�, prod.���, prod.�޸�, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.�׸� = "����"
                prod.��� = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.�޸� = "" Else prod.�޸� = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.������ȣ, prod.�׸�, prod.���, prod.�޸�, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.�׸� = "�ΰǺ�"
                prod.��� = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.�޸� = "" Else prod.�޸� = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.������ȣ, prod.�׸�, prod.���, prod.�޸�, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.�׸� = "��Ÿ"
                prod.��� = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.�޸� = "" Else prod.�޸� = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.������ȣ, prod.�׸�, prod.���, prod.�޸�, regDate
            End If
            
            est.���డ = .Cells(i, 21)
            est.�����ݾ� = .Cells(i, 22)
            est.���� = .Cells(i, 23)
            If est.�����ݾ� = "" Or est.�����ݾ� = "0" Then est.������ = "" Else est.������ = .Cells(i, 24)
            est.���ֱݾ� = .Cells(i, 25)
            est.�������� = .Cells(i, 26)
            est.������� = regDate
            
        End With
        
        '���� ���
        Insert_Record shtEstimateData, est.������ȣ, est.�����ȣ, est.�ŷ�ó, est.�����, est.ǰ��, _
                est.�԰�, est.����, est.����, est.�����ܰ�, est.�����ݾ�, est.������, est.������, est.������, est.��ǰ��, est.������, _
                est.���డ, est.�����ݾ�, est.����, est.������, est.���ֱݾ�, est.��������, est.�������, est.��������
        
        importCount = importCount + 1
        
        '�޸� ���
        For j = 1 To endCol
            If j < 16 Or j > 20 Then
                '��������׸� ���� �ƴ� ��츸 ����
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.������ȣ, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type1 = importCount
End Function

Function Get_Date_Convert(inputStr As String, Y As String)

    Dim pos As Long
    Dim M, D As String
    
    pos = InStr(inputStr, "/")
    If pos <> 0 Then
        M = Left(inputStr, pos - 1)
        D = Right(inputStr, Len(inputStr) - pos)
        Get_Date_Convert = DateSerial(Y, M, D)
    Else
        Get_Date_Convert = inputStr
    End If
End Function
