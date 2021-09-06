VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOperation 
   Caption         =   "UserForm1"
   ClientHeight    =   11505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
   OleObjectBlob   =   "frmOperation.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim operation As Variant
    
    '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            If contr.Name Like "lbl2*" Then
                'contr.BackColor = RGB(48, 84, 150)
                'contr.ForeColor = RGB(255, 255, 255)
            ElseIf contr.Name Like "lbl3*" Then
                contr.BackColor = RGB(221, 235, 247)
            Else
                contr.BackColor = RGB(242, 242, 242)
            End If
        End If
    Next
    
    '�� ��ġ ����
    If productionFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = operationFormX
        Me.top = operationFormY
    End If
    
    'InitializeCboCategory           '�з�
    InitializeLswOperationList    '��� ���
    'InitializeCboProductonUnit  '��������׸� ����
    'InitializeLswOrderCustomerAutoComplete   '���ְŷ�ó �ڵ��ϼ�
    
    'ClearProductionInput
    
End Sub

Sub InitializeLswOperationList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '��� ����� �о��
    db = Get_DB(shtOperation)
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwManual
        .CheckBoxes = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "No", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "��������", 34, lvwColumnCenter
        .ColumnHeaders.Add , , "�з�", 34, , lvwColumnCenter
        .ColumnHeaders.Add , , "�ŷ�ó", 70
        .ColumnHeaders.Add , , "����", 60
        .ColumnHeaders.Add , , "�԰�", 80
        .ColumnHeaders.Add , , "����", 44, lvwColumnRight
        .ColumnHeaders.Add , , "����", 44, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�޸�", 92
        .ColumnHeaders.Add , , "�������", 0
        
        .ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))
                li.ListSubItems.Add , , i
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 5)
                
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 13)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , db(i, 8)
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                
                li.Selected = False
                
                If IsNumeric(db(i, 11)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 11))
                End If
                
                
            Next
            
            Me.txtProductionTotalCost.value = Format(totalCost, "#,##0")
        End If
    End With
End Sub
