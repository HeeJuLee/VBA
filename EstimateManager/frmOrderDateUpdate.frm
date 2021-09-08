VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderDateUpdate 
   Caption         =   "���� �ϰ�����"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15555
   OleObjectBlob   =   "frmOrderDateUpdate.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmOrderDateUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim li As ListItem
    Dim count As Long
    Dim contr As Control
    
    If isFormLoaded("frmEstimateUpdate") = False Then
        MsgBox "���� ����ȭ���� ã�� �� �����ϴ�.", vbInformation, "�۾� Ȯ��"
        Exit Sub
    End If
    
    count = 0
    For Each li In frmEstimateUpdate.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    
    If count = 0 Then
        MsgBox "�ϰ� ������ ���ָ� �����ϼ���.", vbInformation, "�۾� Ȯ��"
        Exit Sub
    End If
    
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
    If orderDateUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderDateUpdateFormX
        Me.top = orderDateUpdateFormY
    End If
    
    InitializeLswOrderList
End Sub

Sub InitializeLswOrderList()
    Dim orgItem, destItem As ListItem
    
     '����Ʈ�� �� ����
    With Me.lswOrderList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "�ŷ�ó", 60
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "����", 83, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 83, lvwColumnCenter
        .ColumnHeaders.Add , , "�԰�", 83, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 83, lvwColumnCenter
        .ColumnHeaders.Add , , "��꼭", 83, lvwColumnCenter
        .ColumnHeaders.Add , , "������", 83, lvwColumnCenter
    
        .ListItems.Clear
        For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
            If orgItem.Selected = True Then
                Set destItem = .ListItems.Add(, , orgItem.Text)   'ID
                destItem.ListSubItems.Add , , orgItem.SubItems(4)       '�ŷ�ó
                destItem.ListSubItems.Add , , orgItem.SubItems(5)       'ǰ��
                destItem.ListSubItems.Add , , orgItem.SubItems(12)       '������
                destItem.ListSubItems.Add , , orgItem.SubItems(13)       '������
                destItem.ListSubItems.Add , , orgItem.SubItems(14)       '�԰���
                destItem.ListSubItems.Add , , orgItem.SubItems(15)      '����
                destItem.ListSubItems.Add , , orgItem.SubItems(16)      '��꼭
                destItem.ListSubItems.Add , , orgItem.SubItems(17)       '������
                destItem.Selected = False
            End If
        Next
    End With
End Sub

Private Sub btnOrderDateSave_Click()
    If chkOrderDate Then
        OrderDateUpdate "��������", Me.txtOrderDate.value
    End If
    If chkDueDate Then
        OrderDateUpdate "��������", Me.txtDueDate.value
    End If
    If chkReceivingDate Then
        OrderDateUpdate "�԰�����", Me.txtReceivingDate.value
    End If
    If chkSpecificationDate Then
        OrderDateUpdate "��������", Me.txtSpecificationDate.value
    End If
    If chkTaxinvoiceDate Then
        OrderDateUpdate "��꼭����", Me.txtTaxinvoiceDate.value
    End If
    If chkPaymentDate Then
        OrderDateUpdate "��������", Me.txtPaymentDate.value
    End If

    Unload Me
    
    frmEstimateUpdate.InitializeLswOrderList
End Sub

Sub OrderDateUpdate(fieldName, value)
    Dim subItemNo, orderColNo, findRow As Long
    Dim orgItem As ListItem
    
    Select Case fieldName
        Case "��������"
            orderColNo = 13  'shtOrder�� �� ��ȣ
            subItemNo = 12  'frmEstimateUpdate orderList�� subitem no
        Case "��������"
            orderColNo = 14
            subItemNo = 13
        Case "�԰�����"
            orderColNo = 15
            subItemNo = 14
        Case "��������"
            orderColNo = 16
            subItemNo = 15
        Case "��꼭����"
            orderColNo = 17
            subItemNo = 16
        Case "��������"
            orderColNo = 18
            subItemNo = 17
    End Select
    
    For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
        If orgItem.Selected = True Then
            'DB ������Ʈ
            Update_Record_Column shtOrder, orgItem.Text, fieldName, value
            'shtOrderAdmin ��Ʈ ������Ʈ
            frmEstimateUpdate.UpdateShtOrderField orgItem.Text, orderColNo, value
        End If
    Next
End Sub

Private Sub btnOrderDateClose_Click()
    Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
    chkOrderDate_Change
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
    chkDueDate_Change
End Sub

Private Sub imgReceivingDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtReceivingDate
    chkReceivingDate_Change
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    chkSpecificationDate_Change
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxinvoiceDate
    chkTaxinvoiceDate_Change
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
    chkPaymentDate_Change
End Sub

Private Sub chkOrderDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkOrderDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(3).Text = Me.txtOrderDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(3).Text = orgItem.ListSubItems(12).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub

Private Sub chkDueDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkDueDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(4).Text = Me.txtDueDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(4).Text = orgItem.ListSubItems(13).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub

Private Sub chkReceivingDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkReceivingDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(5).Text = Me.txtReceivingDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(5).Text = orgItem.ListSubItems(14).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub


Private Sub chkSpecificationDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkSpecificationDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(6).Text = Me.txtSpecificationDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(6).Text = orgItem.ListSubItems(15).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub

Private Sub chkTaxinvoiceDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkTaxinvoiceDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(7).Text = Me.txtTaxinvoiceDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(7).Text = orgItem.ListSubItems(16).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub


Private Sub chkPaymentDate_Change()
    Dim i As Long
    Dim orgItem As ListItem
    
    With Me.lswOrderList
        If chkPaymentDate.value = True Then
            For i = 1 To .ListItems.count
                .ListItems(i).ListSubItems(8).Text = Me.txtPaymentDate.value
            Next
        Else
            i = 1
            For Each orgItem In frmEstimateUpdate.lswOrderList.ListItems
                If orgItem.Selected = True Then
                    .ListItems(i).ListSubItems(8).Text = orgItem.ListSubItems(17).Text
                    i = i + 1
                End If
            Next
        End If
    End With
End Sub

Private Sub UserForm_Layout()
    orderDateUpdateFormX = Me.Left
    orderDateUpdateFormY = Me.top
End Sub

