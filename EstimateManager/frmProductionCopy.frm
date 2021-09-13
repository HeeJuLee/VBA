VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductionCopy 
   Caption         =   "��������׸����� ��������"
   ClientHeight    =   10725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "frmProductionCopy.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmProductionCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim selectedEstimateId As Variant

Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim estimate As Variant
    
    '�� ��ġ ����
    If productionCopyFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = productionCopyFormX
        Me.top = productionCopyFormY
    End If
    
    '�ؽ�Ʈ�ڽ� �� ��ġ ����
    For Each contr In Me.Controls
        If contr.Name Like "Label*" Then
            contr.top = contr.top + 2
        End If
    Next
    
    InitializeLswEstimateList
    InitializeLswOrderList
    
End Sub

Sub InitializeLswEstimateList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '����Ʈ�� �� ����
    With Me.lswEstimateList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwAutomatic
        .CheckBoxes = False
        .SmallIcons = Me.ImageList1
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "������ȣ", 76
        .ColumnHeaders.Add , , "�ŷ�ó", 60
        .ColumnHeaders.Add , , "�����", 50
        .ColumnHeaders.Add , , "������", 230
        .ColumnHeaders.Add , , "��������", 60, lvwColumnCenter
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswEstimateList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    'Ű����� �˻��� ���� ����� ������
    If Me.txtKeyword.value <> "" Then
        db = Get_DB(shtEstimate)
        db = Filtered_DB(db, Me.txtKeyword.value)
        If Not IsEmpty(db) Then
            db = Filtered_DB(db, Me.txtKeyword2.value)
        End If
        
        With Me.lswEstimateList
            .ListItems.Clear
            If Not IsEmpty(db) Then
                For i = 1 To UBound(db)
                    Set li = .ListItems.Add(, , db(i, 1))
                    li.ListSubItems.Add , , db(i, 2)
                    li.ListSubItems.Add , , db(i, 4)
                    li.ListSubItems.Add , , db(i, 5)
                    li.ListSubItems.Add , , db(i, 6)
                    li.ListSubItems.Add , , db(i, 14)
                    
                    li.Selected = False
                Next
            End If
        End With
        
        selectedEstimateId = ""
    End If
End Sub

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '����Ʈ�� �� ����
    With Me.lswOrderList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwAutomatic
        .CheckBoxes = False
        .SmallIcons = Me.ImageList1
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 68
        .ColumnHeaders.Add , , "ǰ��", 100
        .ColumnHeaders.Add , , "����", 50
        .ColumnHeaders.Add , , "�԰�", 50
        .ColumnHeaders.Add , , "����", 44, lvwColumnRight
        .ColumnHeaders.Add , , "����", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 60, lvwColumnRight
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswOrderList(estimateId)
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '����ID�� �ش��ϴ� ���ָ���Ʈ�� �о��
    db = Get_DB(shtOrder)
    db = Filtered_DB(db, estimateId, 28, True)
    
     '����Ʈ�� �� ����
    With Me.lswOrderList
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If db(i, 4) <> "����" Then
                    Set li = .ListItems.Add(, , db(i, 1))
                    li.ListSubItems.Add , , db(i, 4)
                    li.ListSubItems.Add , , db(i, 6)
                    li.ListSubItems.Add , , db(i, 7)
                    li.ListSubItems.Add , , db(i, 8)
                    li.ListSubItems.Add , , db(i, 9)
                    li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                    li.ListSubItems.Add , , db(i, 11)
                    li.ListSubItems.Add , , Format(db(i, 12), "#,##0")
                    li.ListSubItems.Add , , Format(db(i, 13), "#,##0")
                    li.Selected = False
                End If
            Next
            
        End If
    End With
End Sub

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwAutomatic
        .CheckBoxes = False
        .SmallIcons = Me.ImageList1
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 68
        .ColumnHeaders.Add , , "ǰ��", 100
        .ColumnHeaders.Add , , "����", 50
        .ColumnHeaders.Add , , "�԰�", 50
        .ColumnHeaders.Add , , "����", 44, lvwColumnRight
        .ColumnHeaders.Add , , "����", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�޸�", 0
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswProductionList(estimateId)
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, estimateId, 2, True)
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))
                li.ListSubItems.Add , , db(i, 13)
                li.ListSubItems.Add , , db(i, 2)
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 5)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , Format(db(i, 8), "#,##0")
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                
                li.Selected = False
            Next
            
        End If
    End With
End Sub

Private Sub lswEstimateList_Click()
    With Me.lswEstimateList
        If Not .selectedItem Is Nothing Then
            selectedEstimateId = .selectedItem.Text
            SetLswOrderList (selectedEstimateId)
        End If
    End With
End Sub

Private Sub lswEstimateList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswEstimateList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub lswOrderList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswOrderList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub btnEstimateSearch_Click()
    SetLswEstimateList
End Sub

Private Sub btnProductionClose_Click()
    Unload Me
End Sub

Private Sub btnProductionCopyAll_Click()
    ProductionCopy "all"
End Sub

Private Sub btnProductionCopy_Click()
    ProductionCopy ""
End Sub

Sub ProductionCopy(all)
    Dim count As Long
    Dim yn As Variant
    Dim li As ListItem
    
    If selectedEstimateId = "" Then
        MsgBox "������ ������ �����ϼ���.", vbInformation, "�۾� Ȯ��"
        Exit Sub
    End If
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If all = "all" Then
            count = count + 1
        Else
            If li.Selected = True Then count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
        
    yn = MsgBox(count & "�� �׸��� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If yn = vbNo Then Exit Sub
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Or all = "all" Then
            Insert_Record shtProduction, currentEstimateId, currentManagementId, li.SubItems(2), li.SubItems(3), li.SubItems(4), li.SubItems(5), li.SubItems(6), li.SubItems(7), li.SubItems(8), li.SubItems(9), , li.SubItems(1), Date
            count = count + 1
        End If
    Next
    
    If isFormLoaded("frmProductionManager") Then
        frmProductionManager.RefreshProductionTotalCost
    End If
    
    MsgBox count & "�� �׸��� �����Ͽ����ϴ�.", vbInformation, "�۾� Ȯ��"
    
End Sub

Private Sub txtKeyword_AfterUpdate()
    Me.txtKeyword.value = Trim(Me.txtKeyword.value)
    SetLswEstimateList
End Sub

Private Sub txtKeyword2_AfterUpdate()
    Me.txtKeyword2.value = Trim(Me.txtKeyword2.value)
    SetLswEstimateList
End Sub

Private Sub UserForm_Layout()
    productionCopyFormX = Me.Left
    productionCopyFormY = Me.top
End Sub

