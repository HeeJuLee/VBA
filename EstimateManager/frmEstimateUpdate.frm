VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "���� ����"
   ClientHeight    =   13110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19320
   OleObjectBlob   =   "frmEstimateUpdate.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim orgManagementID As Variant
Dim totlalCheckCount As Long
Dim mouseX As Integer
Dim headerIndex As Integer
Dim beforeSelectedItem As ListItem
Dim bInitialIzed As Boolean
Dim currentEditText, currentCboText As Variant




Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim db As Variant
    Dim contr As Control
    Dim acceptedMemo As Variant
    
    If clickEstimateId <> "" Then              '���ְ������� ����Ŭ���� ���
        currentEstimateId = CLng(clickEstimateId)
        clickEstimateId = ""
    Else
        '������ �� ��ȣ
        cRow = Selection.row
    
        '�����Ͱ� �ִ� ���� �ƴ� ���� ����
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).value = "" Then
            MsgBox "������ ���� ���� ���� ������ �� �������� ��ư�� Ŭ���ϼ���.", vbInformation, "�۾� Ȯ��"
            End
        End If
        
        currentEstimateId = shtEstimateAdmin.Cells(cRow, 3)
    End If
    
    bInitialIzed = False
    
     '�ؽ�Ʈ�ڽ� �� ��ġ ����
    For Each contr In Me.Controls
        If contr.Name Like "Label*" Then
            contr.top = contr.top + 2
        End If
    Next
    
    '�� ��ġ ����
    If estimateUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateUpdateFormX
        Me.top = estimateUpdateFormY
    End If
    
    '���� ������ �о����
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If isEmpty(estimate) Or estimate(1) = "" Then
        MsgBox "���� ������ �о�� �� �����ϴ�. ����ID(" & currentEstimateId & ")", vbInformation, "�۾� Ȯ��"
        End
    End If

    Me.txtID.value = estimate(1)    'ID
    Me.txtEstimateName.value = estimate(6)  '������
    Me.txtManagementID.value = estimate(2)    '������ȣ
    currentManagementId = Me.txtManagementID.value
    Me.txtLinkedID.value = estimate(3)  '�����ȣ
    
    Me.txtCustomer = estimate(4)   '�ŷ�ó
    Me.txtManager = estimate(5)   '�����
    
    Me.txtSize.value = estimate(7)  '�԰�
    Me.txtAmount.value = Format(estimate(8), "#,##0")   '����
    InitializeCboUnit
    Me.cboUnit.value = Trim(estimate(9))  '����, ID�� �����Ƿ� ���� value ������ ���õ�
    Me.txtUnitPrice.value = Format(estimate(10), "#,##0")     '�ܰ�
    Me.txtEstimatePrice.value = Format(estimate(11), "#,##0")     '�ݾ�
    
    Me.txtEstimateDate.value = estimate(12)    '��������
    Me.txtBidDate.value = estimate(13)    '��������
    Me.txtAcceptedDate.value = estimate(14)    '��������
    Me.txtDeliveryDate.value = estimate(15)    '��ǰ����
    Me.txtInsuranceDate.value = estimate(16)    '��������
    
    Me.txtProductionTotalCost.value = Format(estimate(17), "#,##0")   '(����)���డ
    Me.txtBidPrice.value = Format(estimate(18), "#,##0")    '������
    Me.txtBidMargin.value = Format(estimate(19), "#,##0")    '(����)����
    Me.txtBidMarginRate.value = Format(estimate(20), "0.0%")    '(����)������
    Me.txtAcceptedPrice.value = Format(estimate(21), "#,##0")    '���ֱݾ�
    Me.txtAcceptedMargin.value = Format(estimate(22), "#,##0")   '��������
    If IsNumeric(Me.txtAcceptedPrice.value) And IsNumeric(Me.txtAcceptedMargin.value) And Me.txtAcceptedPrice <> "" Then
        If CLng(Me.txtAcceptedMargin.value) <> 0 Then
            If CLng(Me.txtAcceptedPrice.value) = 0 Then
                Me.txtAcceptedMarginRate = "0%"
            Else
                Me.txtAcceptedMarginRate = Format(CLng(Me.txtAcceptedMargin) / CLng(Me.txtAcceptedPrice.value), "0.0%")
            End If
        End If
    End If
    
    Me.txtInsertDate.value = estimate(23)    '�������
    Me.txtUpdateDate.value = estimate(24)    '��������
    
    InitializeCboCategory
    Me.cboCategory.value = Trim(estimate(25))   '�з�1
    Me.txtDueDate.value = estimate(26)              '������
    
    Me.txtVAT.value = Format(estimate(31), "#,##0")    '�ΰ���
    Me.txtMemo.value = Trim(estimate(32))     '�����޸�
    Me.chkVAT.value = estimate(33)      '�ΰ��� ���� ����
    
    Me.txtPaid.value = Format(estimate(34), "#,##0")      '�Աݾ�
    Me.txtRemaining.value = Format(estimate(35), "#,##0")      '���Աݾ�
    
    '�������� ��������xls�� ����xls���� �޴��� �ٸ��� ����
    '������ �����޸�� ���ָ޸�� ���� ����� ��.
    '�����޸�� ���ָ޸� �ٸ� ���� ���� �����
    '(�����޸� = �����޸� + ���ָ޸�) �̷��� ���߰� ���� �� �����ʿ��� �����ϰ� �޸� ���� ����
    acceptedMemo = Trim(estimate(36))
    If Me.txtMemo.value <> acceptedMemo Then
        If Me.txtMemo.value = "" Then
            Me.txtMemo.value = acceptedMemo
        Else
            Me.txtMemo.value = Me.txtMemo.value & vbCrLf & acceptedMemo
        End If
    End If
    
    '���� ID (ID_����)
    Me.txtAcceptedID.value = estimate(37)
    currentAcceptedId = estimate(37)
    If Me.txtAcceptedID.value = "" Then
        '����ID�� ������ ���ְ��� ��Ʈ�� unable ��Ŵ
        frmOrder.Visible = False
        btnAcceptedInsert.Visible = True
        frmEstimateUpdate.Height = 260
    Else
        frmOrder.Visible = True
        btnAcceptedInsert.Visible = False
    End If
    
    '���� �� ������ȣ
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '���� ��Ȳ
    InitializeCboOrderCategory  '���� �з�2
    InitializeLswCustomerAutoComplete   '�ŷ�ó �ڵ��ϼ�
    InitializeLswManagerAutoComplete    '����� �ڵ��ϼ�
    InitializeLswPaymentList      '���� ��Ȳ
    InitializeCboEstimatePayMethod '��������-����
    
    Me.txtSize.SetFocus
    
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtEstimateCustomer, True)

    Update_Cbo Me.cboCustomer, db
End Sub


Sub InitializeCboManager()
    Dim db As Variant
    Dim i As Long
    
    '����� DB�� �о�ͼ�
    db = Get_DB(shtEstimateManager, True)
    '�ŷ�ó������ ���͸�
    db = Filtered_DB(db, Me.cboCustomer.value, 1, True)
    
    '���� �޺��ڽ� ���������
    Me.cboManager.Clear
    
    '����ڰ� ������ �޺��ڽ��� �߰���
    If Not isEmpty(db) Then
        Update_Cbo Me.cboManager, db, 2
    End If
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeCboCategory()
    Dim db As Variant
    db = Get_DB(shtEstimateCategory, True)

    Update_Cbo Me.cboCategory, db
End Sub

Sub InitializeCboOrderCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboOrderCategory, db
End Sub

Sub InitializeCboEstimatePayMethod()
    Dim db As Variant
    db = Get_DB(shtEstimatePayMethod, True)

    Update_Cbo Me.cboEstimatePayMethod, db
End Sub

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j As Long
    Dim totalCost As Double
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� ���� ������ �о��
    db = Get_DB(shtOrder)
    If Not isEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.value, 28, True)
    End If
    If Not isEmpty(db) Then
        db = Filtered_DB(db, "<>" & "����", 4)
    End If
    
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
        .LabelEdit = lvwManual
        .SmallIcons = Me.ImageList1
        .Sorted = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 50
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "����", 60
        .ColumnHeaders.Add , , "�԰�", 62
        .ColumnHeaders.Add , , "����", 30, lvwColumnRight
        .ColumnHeaders.Add , , "����", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "�԰�", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "��꼭", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 30
        
        '.ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not isEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))   'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_����
                li.ListSubItems.Add , , db(i, 5)        '������ȣ
                li.ListSubItems.Add , , db(i, 4)        '�з�
                li.ListSubItems.Add , , db(i, 6)        '�ŷ�ó
                li.ListSubItems.Add , , db(i, 7)        'ǰ��
                li.ListSubItems.Add , , db(i, 8)        '����
                li.ListSubItems.Add , , db(i, 9)        '�԰�
                li.ListSubItems.Add , , db(i, 10)        '����
                li.ListSubItems.Add , , db(i, 11)       '����
                li.ListSubItems.Add , , Format(db(i, 12), "#,##0")      '�ܰ�
                li.ListSubItems.Add , , Format(db(i, 13), "#,##0")      '�ݾ�
                li.ListSubItems.Add , , db(i, 16)       '������
                li.ListSubItems.Add , , db(i, 17)       '������
                li.ListSubItems.Add , , db(i, 18)       '�԰���
                li.ListSubItems.Add , , db(i, 20)       '����
                li.ListSubItems.Add , , db(i, 21)       '��꼭
                li.ListSubItems.Add , , db(i, 22)       '����
                li.ListSubItems.Add , , "����"       '����
                li.Selected = False
                
                If IsNumeric(db(i, 13)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 13))
                End If
            Next
        End If
        
        If totalCost <> 0 Then
            Me.txtExecutionCost.value = Format(totalCost, "#,##0")
        End If
    End With
End Sub

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeLswManagerAutoComplete()
    
    With Me.lswManagerAutoComplete
        .View = lvwList
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeLswPaymentList()
    Dim db As Variant
    Dim i, j, totalPaid As Long
    Dim li As ListItem

    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With

     '����Ʈ�� �� ����
    With Me.lswPaymentList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwManual
        .SmallIcons = Me.ImageList1
        .Sorted = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "��꼭", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "������", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "�Աݾ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�Աݿ�����", 70, lvwColumnRight
        .ColumnHeaders.Add , , "��������", 60, lvwColumnCenter
        .ColumnHeaders.Add , , "�޸�", 110
        .ColumnHeaders.Add , , "�ΰ���", 60, lvwColumnRight
        .ColumnHeaders.Add , , "����������", 0

        .ListItems.Clear
        
        '����ID�� �ش��ϴ� ���� �̷¸� �о��
        db = Get_DB(shtPayment)
        If Not isEmpty(db) Then
            db = Filtered_DB(db, Me.txtID.value, 2, True)
        End If

        totalPaid = 0
        If Not isEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))   'ID
                li.ListSubItems.Add , , db(i, 2)       'ID_����
                li.ListSubItems.Add , , db(i, 3)        '������ȣ
                li.ListSubItems.Add , , db(i, 4)        '����
                li.ListSubItems.Add , , db(i, 5)        '��꼭
                li.ListSubItems.Add , , db(i, 6)        '����
                li.ListSubItems.Add , , Format(db(i, 7), "mm" & "��")      '������
                li.ListSubItems.Add , , Format(db(i, 8), "#,##0")        '�Աݾ�
                li.ListSubItems.Add , , Format(db(i, 9), "#,##0")        '�Աݿ�����
                li.ListSubItems.Add , , db(i, 10)        '��������
                li.ListSubItems.Add , , db(i, 11)       '�޸�
                li.ListSubItems.Add , , Format(db(i, 12), "#,##0")       '�ΰ���
                li.ListSubItems.Add , , db(i, 7)       '����������
                li.Selected = False

                If IsNumeric(db(i, 8)) Then
                    '�Ա� �հ� ����
                    totalPaid = totalPaid + CLng(db(i, 8))
                End If
            Next
        End If

        Me.txtPaid.value = Format(totalPaid, "#,##0")
        CalculatePayment
        
    End With

End Sub

Sub UpdateEstimateOrderValue(fieldName, fieldValue)
        
    '����DB ����
    Update_Record_Column shtEstimate, currentEstimateId, fieldName, fieldValue
    
    '������Ʈ ����
    UpdateShtEstimateField currentEstimateId, fieldName, fieldValue
    
    '����DB ����
    Update_Record_Column shtOrder, currentAcceptedId, fieldName, fieldValue
    
    '���ֽ�Ʈ ����
    UpdateShtOrderField currentAcceptedId, fieldName, fieldValue
    
End Sub

Sub UpdateEstimateValue(fieldName, fieldValue)
        
    '����DB ����
    Update_Record_Column shtEstimate, currentEstimateId, fieldName, fieldValue
    
    '������Ʈ ����
    UpdateShtEstimateField currentEstimateId, fieldName, fieldValue
End Sub

Sub UpdateOrderValue(fieldName, fieldValue)
    '����DB ����
    Update_Record_Column shtOrder, currentAcceptedId, fieldName, fieldValue
    
    '���ֽ�Ʈ ����
    UpdateShtOrderField currentAcceptedId, fieldName, fieldValue
    
End Sub

Sub InsertAccepted()

    '���ֹ��� ���̺� ���� ���
    Insert_Record shtOrder, _
            , Me.cboCategory.value, "����", Me.txtManagementID.value, _
            Me.txtCustomer.value, _
            Me.txtEstimateName.value, _
            Me.txtManager.value, _
            Me.txtSize.value, _
            Me.txtAmount.value, _
            Me.cboUnit.value, _
            , _
            , _
            , _
            Date, , , , , _
            , , , , _
            , , _
            Date, , _
            CLng(Me.txtID.value), , False, "����"

    '����� ����ID�� ���� ���̺� ������Ʈ, �������ڴ� ����
    Update_Record_Column shtEstimate, Me.txtID, "ID_����", Get_LastID(shtOrder)
    Update_Record_Column shtEstimate, Me.txtID, "����", Date
    
    '���� ���� ���
    Unload frmEstimateUpdate
    
    clickEstimateId = Me.txtID
    frmEstimateUpdate.Show (False)
    
End Sub

'���� ����Ʈ�� �� DB ����
Sub UpdateOrderListValue(id, headerIndex, value)
    Dim fieldName As String

    Select Case headerIndex
        Case 4  '�з�
            fieldName = "�з�2"
        Case 5  '�ŷ�ó
            fieldName = "�ŷ�ó"
        Case 6  'ǰ��
            fieldName = "ǰ��"
        Case 7  '����
            fieldName = "����"
        Case 8  '�԰�
            fieldName = "�԰�"
        Case 9  '����
            fieldName = "����"
        Case 10  '����
            fieldName = "����"
        Case 11  '�ܰ�
            fieldName = "�ܰ�"
        Case 12  '�ݾ�
            fieldName = "�ݾ�"
        Case 13  '����
            fieldName = "����"
        Case 14  '����
            fieldName = "����"
        Case 15  '�԰�
            fieldName = "�԰�"
        Case 16  '����
            fieldName = "����"
        Case 17  '��꼭
            fieldName = "��꼭"
        Case 18  '����
            fieldName = "����"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtOrder, id, fieldName, value
        Update_Record_Column shtOrder, id, "��������", Date
        UpdateShtOrderField id, fieldName, value
    End If

End Sub

Sub UpdatePaymentListValue(id, headerIndex, value)
    Dim fieldName As String

    Select Case headerIndex
        Case 4
            fieldName = "����"
        Case 5
            fieldName = "��꼭"
        Case 6
            fieldName = "����"
        Case 7
            fieldName = "������"
        Case 8
            fieldName = "�Աݾ�"
        Case 9
            fieldName = "�Աݿ�����"
        Case 10
            fieldName = "��������"
        Case 11
            fieldName = "�޸�"
        Case 12
            fieldName = "�ΰ���"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtPayment, id, fieldName, value
        Update_Record_Column shtPayment, id, "��������", Date
    End If
    
    Select Case headerIndex
        Case 4, 5, 6, 7, 10, 12
            '���� ���� ������ ���� �� ���������̸� ����/��꼭/����/������ �����͸� ����DB/����DB/����������Ʈ/���ְ�����Ʈ�� ����
            If lswPaymentList.ListItems(lswPaymentList.ListItems.count).Selected = True Then
                UpdateEstimateOrderValue fieldName, value
            End If
    End Select

End Sub

Sub SelectOrderListColumn()
    Dim ItemSel    As ListItem
    
    If Not lswOrderList.selectedItem Is Nothing Then
        If headerIndex = lswOrderList.ColumnHeaders.count Then
            frmEdit.Visible = False
            txtEdit.Visible = False
            cboOrderCategory.Visible = False
        End If
        
        Set ItemSel = lswOrderList.selectedItem
        ItemSel.EnsureVisible
            
        If headerIndex > 4 And headerIndex < lswOrderList.ColumnHeaders.count Then
            With frmEdit
                .Visible = True
                .top = ItemSel.top + lswOrderList.top
                .Left = lswOrderList.ColumnHeaders(headerIndex).Left + lswOrderList.Left
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.txtEdit
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width
                .Height = lswOrderList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentEditText = .Text
            End With
            Me.cboOrderCategory.Visible = False
        ElseIf headerIndex = 4 Then
            With frmEdit
                .Visible = True
                .top = ItemSel.top + lswOrderList.top
                .Left = lswOrderList.ColumnHeaders(headerIndex).Left + lswOrderList.Left
                .Width = cboOrderCategory.Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With cboOrderCategory
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Height = lswOrderList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentCboText = .Text
            End With
            Me.txtEdit.Visible = False
        End If
    End If

End Sub

Sub SelectPaymentListColumn()
    Dim ItemSel    As ListItem
    
    If Not lswPaymentList.selectedItem Is Nothing Then
        If headerIndex = lswPaymentList.ColumnHeaders.count Then
            frmPaymentEdit.Visible = False
            txtPaymentEdit.Visible = False
            cboEstimatePayMethod.Visible = False
        End If
        
        Set ItemSel = lswPaymentList.selectedItem
        ItemSel.EnsureVisible
        
        If headerIndex = 10 Then
            With frmPaymentEdit
                .Visible = True
                .top = ItemSel.top + lswPaymentList.top
                .Left = lswPaymentList.ColumnHeaders(headerIndex).Left + lswPaymentList.Left
                .Width = Me.cboEstimatePayMethod.Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.cboEstimatePayMethod
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Height = lswPaymentList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentCboText = .Text
            End With
            Me.txtPaymentEdit.Visible = False
        ElseIf headerIndex > 3 And headerIndex < lswPaymentList.ColumnHeaders.count Then
            With frmPaymentEdit
                .Visible = True
                .top = ItemSel.top + lswPaymentList.top
                .Left = lswPaymentList.ColumnHeaders(headerIndex).Left + lswPaymentList.Left
                .Width = lswPaymentList.ColumnHeaders(headerIndex).Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.txtPaymentEdit
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Width = lswPaymentList.ColumnHeaders(headerIndex).Width
                .Height = lswPaymentList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentEditText = .Text
            End With
            Me.cboEstimatePayMethod.Visible = False
        End If
    End If

End Sub

Sub DeleteOrderList()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "������ ���ָ� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    yn = MsgBox("������ " & count & "�� ���ָ� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If yn = vbNo Then Exit Sub

    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            '���� ���̺��� ����
            Delete_Record shtOrder, li.Text
        End If
    Next
    
    If count > 0 Then
        InitializeLswOrderList
    End If
End Sub
'
'Sub UpdateShtEstimate(estimateId)
'    Dim findRow As Long
'
'    findRow = isExistInSheet(shtEstimateAdmin.Range("B6"), estimateId)
'    If findRow <> 0 Then
'        shtEstimateAdmin.Cells(findRow, 4).value = Me.txtManagementID.value
'        shtEstimateAdmin.Cells(findRow, 5).value = Me.txtCustomer.value
'        shtEstimateAdmin.Cells(findRow, 6).value = Me.txtManager.value
'        shtEstimateAdmin.Cells(findRow, 7).value = Me.cboCategory.value
'        shtEstimateAdmin.Cells(findRow, 8).value = Me.txtEstimateName.value
'        shtEstimateAdmin.Cells(findRow, 9).value = Me.txtSize.value
'        shtEstimateAdmin.Cells(findRow, 10).value = Me.txtAmount.value
'        shtEstimateAdmin.Cells(findRow, 11).value = Me.cboUnit.value
'        shtEstimateAdmin.Cells(findRow, 12).value = Me.txtUnitPrice.value
'        shtEstimateAdmin.Cells(findRow, 13).value = Me.txtEstimatePrice.value
'        shtEstimateAdmin.Cells(findRow, 14).value = Me.txtEstimateDate.value
'        shtEstimateAdmin.Cells(findRow, 15).value = Me.txtBidDate.value
'        shtEstimateAdmin.Cells(findRow, 16).value = Me.txtAcceptedDate.value
'        shtEstimateAdmin.Cells(findRow, 17).value = Me.txtDueDate.value
'        shtEstimateAdmin.Cells(findRow, 18).value = Me.txtDeliveryDate.value
'        shtEstimateAdmin.Cells(findRow, 19).value = Me.txtInsuranceDate.value
'        shtEstimateAdmin.Cells(findRow, 20).value = Me.txtProductionTotalCost.value
'        shtEstimateAdmin.Cells(findRow, 21).value = Me.txtBidPrice.value
'        shtEstimateAdmin.Cells(findRow, 22).value = Me.txtBidMargin.value
'        shtEstimateAdmin.Cells(findRow, 23).value = Me.txtBidMarginRate.value
'        shtEstimateAdmin.Cells(findRow, 24).value = Me.txtAcceptedPrice.value
'        shtEstimateAdmin.Cells(findRow, 25).value = Me.txtAcceptedMargin.value
'        shtEstimateAdmin.Cells(findRow, 26).value = Me.txtSpecificationDate.value
'        shtEstimateAdmin.Cells(findRow, 27).value = Me.txtTaxinvoiceDate.value
'        shtEstimateAdmin.Cells(findRow, 28).value = Me.txtPaymentDate.value
'        shtEstimateAdmin.Cells(findRow, 29).value = Me.txtExpectPaymentDate.value
'        shtEstimateAdmin.Cells(findRow, 30).value = Me.txtVAT.value
'        shtEstimateAdmin.Cells(findRow, 31).value = Me.txtInsertDate.value
'        shtEstimateAdmin.Cells(findRow, 32).value = Date
'    End If
'End Sub

Sub UpdateShtEstimateField(estimateId, fieldName, value)
    Dim findRow As Long
    Dim colNo As Long
    
    findRow = isExistInSheet(shtEstimateAdmin.Range("C6"), estimateId)
    If findRow > 0 Then
        colNo = 0
        Select Case fieldName
            Case "������ȣ"
                colNo = 4
            Case "�ŷ�ó"
                colNo = 5
            Case "�����"
                colNo = 6
            Case "�з�1"
                colNo = 7
            Case "������"
                colNo = 8
            Case "�԰�"
                colNo = 9
            Case "����"
                colNo = 10
            Case "����"
                colNo = 11
            Case "�ܰ�"
                colNo = 12
            Case "�ݾ�"
                colNo = 13
            Case "����"
                colNo = 14
            Case "����"
                colNo = 15
            Case "����"
                colNo = 16
            Case "����"
                colNo = 17
            Case "��ǰ"
                colNo = 18
            Case "����"
                colNo = 19
            Case "���డ(����)"
                colNo = 20
            Case "�����ݾ�"
                colNo = 21
            Case "����(����)"
                colNo = 22
            Case "������(����)"
                colNo = 23
            Case "���ֱݾ�"
                colNo = 24
            Case "��������"
                colNo = 25
            Case "����"
                colNo = 26
            Case "��꼭"
                colNo = 27
            Case "����"
                colNo = 28
            Case "������"
                colNo = 29
            Case "�ΰ���"
                colNo = 30
            Case "�������"
                colNo = 31
            Case "��������"
                colNo = 32
        End Select
      
        If colNo <> 0 Then
            shtEstimateAdmin.Cells(findRow, colNo).value = value
        End If
    End If
    
End Sub

'Sub UpdateShtOrder(orderId)
'    Dim findRow As Long
'
'    findRow = isExistInSheet(shtOrderAdmin.Range("C6"), orderId)
'    If findRow <> 0 Then
'        shtOrderAdmin.Cells(findRow, 5).value = Me.txtManagementID.value
'        shtOrderAdmin.Cells(findRow, 6).value = Me.cboCategory.value
'        shtOrderAdmin.Cells(findRow, 8).value = Me.txtCustomer.value
'        shtOrderAdmin.Cells(findRow, 9).value = Me.txtEstimateName.value
'        shtOrderAdmin.Cells(findRow, 10).value = Me.txtManager.value
'        shtOrderAdmin.Cells(findRow, 11).value = Me.txtSize.value
'        shtOrderAdmin.Cells(findRow, 12).value = Me.txtAmount.value
'        shtOrderAdmin.Cells(findRow, 13).value = Me.cboUnit.value
'        shtOrderAdmin.Cells(findRow, 14).value = Me.txtUnitPrice.value
'        shtOrderAdmin.Cells(findRow, 15).value = Me.txtEstimatePrice.value
'        shtOrderAdmin.Cells(findRow, 17).value = Me.txtAcceptedDate.value
'        shtOrderAdmin.Cells(findRow, 19).value = Me.txtDueDate.value
'        shtOrderAdmin.Cells(findRow, 21).value = Me.txtDeliveryDate.value
'        shtOrderAdmin.Cells(findRow, 22).value = Me.txtSpecificationDate.value
'        shtOrderAdmin.Cells(findRow, 23).value = Me.txtTaxinvoiceDate.value
'        shtOrderAdmin.Cells(findRow, 24).value = Me.txtPaymentDate.value
'        shtOrderAdmin.Cells(findRow, 25).value = Me.txtExpectPaymentDate.value
'        shtOrderAdmin.Cells(findRow, 27).value = Me.txtVAT.value
'        shtOrderAdmin.Cells(findRow, 28).value = Me.txtInsertDate.value
'        shtOrderAdmin.Cells(findRow, 29).value = Date
'    End If
'End Sub

Sub UpdateShtOrderHeaderIndex(orderId, headerIndex, value)
    Dim findRow, fieldNo As Long
    
    findRow = isExistInSheet(shtOrderAdmin.Range("B6"), orderId)
    If findRow <> 0 Then
        Select Case headerIndex
            Case 4  '�з�
                fieldNo = 7
            Case 5  '�ŷ�ó
                fieldNo = 8
            Case 6  'ǰ��
                fieldNo = 9
            Case 7  '����
                fieldNo = 10
            Case 8  '�԰�
                fieldNo = 11
            Case 9  '����
                fieldNo = 12
            Case 10  '����
                fieldNo = 13
            Case 11  '�ܰ�
                fieldNo = 14
            Case 12  '�ݾ�
                fieldNo = 15
            Case 13  '����
                fieldNo = 18
            Case 14  '����
                fieldNo = 19
            Case 15  '�԰�
                fieldNo = 20
            Case 16  '����
                fieldNo = 22
            Case 17  '��꼭
                fieldNo = 23
            Case 18  '����
                fieldNo = 24
        End Select
        
        shtOrderAdmin.Cells(findRow, fieldNo).value = value
    End If
End Sub

Function CheckEstimateUpdateValidation()
    
    '�������� �ԷµǾ����� üũ
    If Me.txtEstimateName.value = "" Then
        MsgBox "�������� �Է��ϼ���.", vbExclamation
        CheckEstimateUpdateValidation = False
        Exit Function
    End If
    
    '������ȣ�� �ԷµǾ����� üũ
    If Me.txtManagementID.value = "" Then
        MsgBox "������ȣ�� �Է��ϼ���.", vbExclamation
        CheckEstimateUpdateValidation = False
        Exit Function
    End If
    
    CheckEstimateUpdateValidation = True
End Function


Sub CalculateEstimatePrice()
    '�ݾ� ���
    '�������� �����̸� �ݾ��� �ܰ�
    If Me.txtUnitPrice = "" Then
        Me.txtEstimatePrice.value = ""
    Else
        If Me.txtAmount.value = "" Then
            Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Else
            Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")
End Sub

Sub CalculateBidMargin()
    '�������װ� �������� ���
    If Me.txtBidPrice.value <> "" And Me.txtProductionTotalCost.value <> "" Then
        '�������� = ������ - ������డ
        Me.txtBidMargin.value = Format(CLng(Me.txtBidPrice.value) - CLng(Me.txtProductionTotalCost.value), "#,##0")
        '�������� = �������� / ������
        If Me.txtBidPrice.value <> "0" Then
            Me.txtBidMarginRate.value = Format(CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value), "0.0%")
        End If
    Else
        Me.txtBidMargin.value = ""
        Me.txtBidMarginRate.value = ""
    End If
End Sub

Sub CalculateAcceptedMargin()

    '��������, ������ ���
    If Me.txtAcceptedPrice.value <> "" And Me.txtExecutionCost.value <> "" Then
        '�������� = ���ֱݾ� - ���డ
        Me.txtAcceptedMargin.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtExecutionCost.value), "#,##0")
        '������ = �������� / ���ֱݾ�
        If Me.txtAcceptedPrice.value <> "0" Then
            Me.txtAcceptedMarginRate.value = Format(CLng(Me.txtAcceptedMargin.value) / CLng(Me.txtAcceptedPrice.value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.value = ""
        Me.txtAcceptedMarginRate.value = ""
    End If
End Sub

Sub CalculatePayment()
   '���Աݾ�, �ΰ��� ���
    If Me.txtAcceptedPrice.value = "" Then
        Me.txtRemaining.value = ""
        Me.txtVAT.value = ""
    Else
        If IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtRemaining.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtPaid.value), "#,##0")
            If Me.chkVAT.value = True Then
                Me.txtVAT.value = 0
            Else
                Me.txtVAT = Format(CLng(Me.txtPaid.value) * 0.1, "#,##0")
            End If
        End If
    End If
End Sub

Sub CalculateEstimateUpdateCost_2()

    '�ݾ� ���
    '�������� �����̸� �ݾ��� �ܰ�
    If Me.txtUnitPrice = "" Then
        Me.txtEstimatePrice.value = ""
    Else
        If Me.txtAmount.value = "" Then
            Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Else
            Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")

    '�������װ� �������� ���
    If Me.txtBidPrice.value <> "" And Me.txtProductionTotalCost.value <> "" Then
        '�������� = ������ - ������డ
        Me.txtBidMargin.value = Format(CLng(Me.txtBidPrice.value) - CLng(Me.txtProductionTotalCost.value), "#,##0")
        '�������� = �������� / ������
        If Me.txtBidPrice.value <> "0" Then
            Me.txtBidMarginRate.value = Format(CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value), "0.0%")
        End If
    Else
        Me.txtBidMargin.value = ""
        Me.txtBidMarginRate.value = ""
    End If

    '��������, ������ ���
    If Me.txtAcceptedPrice.value <> "" And Me.txtExecutionCost.value <> "" Then
        '�������� = ���ֱݾ� - ���డ
        Me.txtAcceptedMargin.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtExecutionCost.value), "#,##0")
        '������ = �������� / ���ֱݾ�
        If Me.txtAcceptedPrice.value <> "0" Then
            Me.txtAcceptedMarginRate.value = Format(CLng(Me.txtAcceptedMargin.value) / CLng(Me.txtAcceptedPrice.value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.value = ""
        Me.txtAcceptedMarginRate.value = ""
    End If

    '�ΰ��� ���
    '���ݰ�꼭 ���ڰ� ���� ���, �ΰ��� ������ ��� �ΰ����� 0
    If Me.txtTaxinvoiceDate.value = "" Or chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        '�ΰ����� ���ֱݾ��� 10%
        If Me.txtAcceptedPrice.value <> "" And Me.txtAcceptedPrice.value <> 0 Then
            Me.txtVAT.value = CLng(Me.txtAcceptedPrice.value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.value, "#,##0")
        End If
    End If

'    '�Աݿ���� ���
'    If Me.txtTaxInvoiceDate.Value = "" Then
'        '���ݰ�꼭 ���ڰ� ���� ���� ���ֱݾ�
'        Me.txtExpectPay.Value = Me.txtAcceptedPrice
'    Else
'        '���ݰ�꼭 ���ڰ� �ִ� ���� ���ֱݾ�+�ΰ���
'        If Me.txtAcceptedPrice.Value <> "" Then
'            Me.txtExpectPay.Value = CLng(Me.txtAcceptedPrice.Value) + CLng(Me.txtVAT.Value)
'        End If
'    End If
'    Me.txtExpectPay.Text = Format(Me.txtExpectPay.Value, "#,##0")
'
'    '�Աݾ� ���
'    If Me.txtPaymentDate.Value = "" Then
'        Me.txtPaid.Value = 0
'    Else
'        Me.txtPaid.Value = Me.txtExpectPay.Value
'        Me.txtPaid.Text = Format(Me.txtPaid.Value, "#,##0")
'    End If
'
'    '���Աݾ� ���
'    Me.txtUnpaid.Value = CLng(Me.txtExpectPay.Value) - CLng(Me.txtPaid.Value)
'    Me.txtUnpaid.Text = Format(Me.txtUnpaid.Value, "#,##0")
    
End Sub

Function CalculateOrderListTotalCost() As Long
    Dim i As Long
    Dim cost, totalCost As Long
    
    With Me.lswOrderList
        For i = 1 To .ListItems.count

            If Not IsNumeric(.ListItems(i).SubItems(11)) Then
                If .ListItems(i).SubItems(11) <> "" Then
                    MsgBox "�ݾ� �ʵ忡 ���ڰ� �ƴ� ���� �־ ���డ �հ踦 ���� �� �����ϴ�.", vbExclamation
                    CalculateOrderListTotalCost = 0
                    Exit Function
                End If
            Else
                totalCost = totalCost + .ListItems(i).SubItems(11)
            End If
        Next
    End With
    
    CalculateOrderListTotalCost = totalCost
End Function

Function CalculatePaymentListTotalCost() As Long
    Dim i As Long
    Dim cost, totalCost As Long
    
    With Me.lswPaymentList
        For i = 1 To .ListItems.count

            If Not IsNumeric(.ListItems(i).SubItems(7)) Then
                If .ListItems(i).SubItems(7) <> "" Then
                    MsgBox "�ݾ� �ʵ忡 ���ڰ� �ƴ� ���� �־ �Աݾ� �հ踦 ���� �� �����ϴ�.", vbExclamation
                    CalculatePaymentListTotalCost = 0
                    Exit Function
                End If
            Else
                totalCost = totalCost + .ListItems(i).SubItems(7)
            End If
        Next
    End With
    
    CalculatePaymentListTotalCost = totalCost
End Function

Function CalculateOrderListPrice(selectedItem As ListItem) As Long
    Dim amount, unitPrice As Variant
    Dim orderPrice As Long

    '����, �ܰ��� ���ϴ� ��쿡�� �ݾ� ����ؼ� �����ؾ� ��
    amount = selectedItem.ListSubItems(8).Text
    unitPrice = selectedItem.ListSubItems(10).Text
    
    If amount = "" Then
        If IsNumeric(unitPrice) Then
            orderPrice = unitPrice
        End If
    ElseIf IsNumeric(amount) And IsNumeric(unitPrice) Then
        orderPrice = amount * unitPrice
    End If
    
    CalculateOrderListPrice = orderPrice
End Function

Function CalculatePaymentListVAT(selectedItem As ListItem) As Long
    Dim paid As Variant
    Dim VAT As Long

    '��꼭 ���� ������ 0
    If selectedItem.ListSubItems(4).Text = "" Then
        CalculatePaymentListVAT = 0
        Exit Function
    End If
    
    '�Աݾ� ���ϴ� ��쿡 �ΰ��� �����ؾ� ��
    paid = selectedItem.ListSubItems(7).Text
    
    If paid = "" Then
        '�Աݾ��� ���� ��쿡 �Աݿ��������� ��
        paid = selectedItem.ListSubItems(8).Text
    End If
    
    If IsNumeric(paid) Then
        VAT = paid * 0.1
    Else
        VAT = 0
    End If
    
    CalculatePaymentListVAT = VAT
End Function

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.value, 2, True)
    
    'DB�� ���� ���� ���
    totalCost = 0
    If Not isEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 11)) Then
                '��� �հ� ����
                totalCost = totalCost + CLng(db(i, 11))
            End If
        Next
    End If
        
    GetProductionTotalCost = totalCost
End Function

Sub ConvertOrderListFormat(textBox, headerIndex)
    Dim value As Variant
    Dim pos As Long
    Dim Y, M, D As Long
    
    value = Trim(textBox.Text)
    
    Select Case headerIndex
        Case 9, 11, 12  '����, �ܰ�, �ݾ� - 1000�ڸ� �޸�
            If IsNumeric(value) Then
                textBox.Text = Format(value, "#,##0")
            End If
        Case 13, 14, 15, 16, 17, 18  '����, ����, �԰�, ����, ��꼭, ���� - ��¥ ��ȯ
            textBox.Text = ConvertDateFormat(textBox.Text)
    End Select
    
End Sub

Sub ConvertPaymentListFormat(textBox, headerIndex)
    Select Case headerIndex
        Case 8, 9, 12  '�Աݾ�, �ΰ��� - 1000�ڸ� �޸�
            If IsNumeric(textBox.Text) Then
                textBox.Text = Format(textBox.Text, "#,##0")
            End If
        Case 4, 5, 6, 7  '����, ��꼭, ����, ������ - ��¥ ��ȯ
            textBox.Text = ConvertDateFormat(textBox.Text)
    End Select
    
End Sub




Private Sub lswOrderList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * x
End Sub

Private Sub lswPaymentList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * x
End Sub


Private Sub btnOrderListInsert_Click()
    Dim lastId As Long
    Dim li As ListItem
    
    '���ָ���Ʈ�信 ���� �߰�
    Insert_Record shtOrder, _
                , , "����", currentManagementId, , , , , , , , , , _
                , , , , , _
                , , , , _
                , , _
                Date, , currentEstimateId, , False
    lastId = Get_LastID(shtOrder)
    
    With Me.lswOrderList
        Set li = .ListItems.Add(, , lastId)   'ID
        li.ListSubItems.Add , , currentEstimateId       'ID_����
        li.ListSubItems.Add , , currentManagementId        '������ȣ
        li.ListSubItems.Add , , "����"        '�з�
        li.ListSubItems.Add , , ""        '�ŷ�ó
        li.ListSubItems.Add , , ""        'ǰ��
        li.ListSubItems.Add , , ""        '����
        li.ListSubItems.Add , , ""        '�԰�
        li.ListSubItems.Add , , ""        '����
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , ""          '�ܰ�
        li.ListSubItems.Add , , ""      '�ݾ�
        li.ListSubItems.Add , , ""       '������
        li.ListSubItems.Add , , ""       '������
        li.ListSubItems.Add , , ""       '�԰���
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , ""       '��꼭
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , "����"       '����
        
        .selectedItem.Selected = False
        li.Selected = True
        li.EnsureVisible
        
        headerIndex = 4
        SelectOrderListColumn
    End With
End Sub

Private Sub btnOrderListBatchUpdate_Click()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "�ϰ� ������ ���ָ� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    frmOrderDateUpdate.Show
End Sub

Private Sub btnPaymentListInsert_Click()
    Dim lastId As Long
    Dim li As ListItem
    
    frmPaymentEdit.Visible = False
    cboEstimatePayMethod.Visible = False
    
    '�����̷¿� ���� �߰�
    Insert_Record shtPayment, _
                        currentEstimateId, currentManagementId, _
                        Date, , , , , , , , , Date, ""
                        
    lastId = Get_LastID(shtPayment)
    
    With Me.lswPaymentList
        Set li = .ListItems.Add(, , lastId)   'ID
        li.ListSubItems.Add , , currentEstimateId       'ID_����
        li.ListSubItems.Add , , currentManagementId        '������ȣ
        li.ListSubItems.Add , , Date        '����
        li.ListSubItems.Add , , ""        '��꼭
        li.ListSubItems.Add , , ""        '����
        li.ListSubItems.Add , , ""       '������
        li.ListSubItems.Add , , ""        '�Աݾ�
        li.ListSubItems.Add , , ""        '�Աݿ�����
        li.ListSubItems.Add , , ""       '��������
        li.ListSubItems.Add , , ""       '�޸�
        li.ListSubItems.Add , , ""      '�ΰ���
        li.ListSubItems.Add , , ""      '����������(��¥����)
        
        .selectedItem.Selected = False
        li.Selected = True
        li.EnsureVisible
        
        headerIndex = 4
        SelectPaymentListColumn
    End With
    
    'DB�� ��Ʈ�� ����/��꼭/����/������ �� ����
    UpdateEstimateOrderValue "����", Date
    UpdateEstimateOrderValue "��꼭", ""
    UpdateEstimateOrderValue "����", ""
    UpdateEstimateOrderValue "������", ""
    UpdateEstimateOrderValue "��������", ""
    
End Sub

Private Sub btnPaymentListDelete_Click()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    Dim spec, tax, paid, month, method As Variant
    
    count = 0
    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "������ �����̷��� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    yn = MsgBox("������ " & count & "�� �̷��� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If yn = vbNo Then Exit Sub

    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            '�̷� ���̺��� ����
            Delete_Record shtPayment, li.Text
        End If
    Next
    
    If count > 0 Then
        InitializeLswPaymentList
    End If
    
    '���� �÷��� �����ؼ� �����̷��� �߰����� �ʵ��� ��
    bDeleteFlag = True
    
    '�� ������ �̷� �������� ����/��꼭/����/�������� ����DB�� ����
    With Me.lswPaymentList
        If .ListItems.count = 0 Then
            spec = ""
            tax = ""
            paid = ""
            month = ""
            method = ""
        Else
            spec = .ListItems(.ListItems.count).SubItems(3)
            tax = .ListItems(.ListItems.count).SubItems(4)
            paid = .ListItems(.ListItems.count).SubItems(5)
            month = .ListItems(.ListItems.count).SubItems(6)
            method = .ListItems(.ListItems.count).SubItems(8)
        End If
        UpdateEstimateOrderValue "����", spec
        UpdateEstimateOrderValue "��꼭", tax
        UpdateEstimateOrderValue "����", paid
        UpdateEstimateOrderValue "������", month
        UpdateEstimateOrderValue "��������", method
    End With
    
    bDeleteFlag = False
End Sub


Private Sub lswOrderList_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.value = ""
    Me.cboOrderCategory.value = ""
End Sub

Private Sub lswPaymentList_Click()
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.value = ""
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

Private Sub lswPaymentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswPaymentList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub btnProduction_Click()
    If isFormLoaded("frmProductionManager") Then
        Unload frmProductionManager
    End If
    frmProductionManager.Show (False)
End Sub


Private Sub btnAcceptedInsert_Click()
    InsertAccepted
End Sub

Private Sub btnPayment_Click()
    frmPaymentManager.Show
End Sub

Private Sub chkDividePay_Click()
    If chkDividePay.value = True Then
        Me.btnPayment.Enabled = True
    Else
        Me.btnPayment.Enabled = False
    End If
End Sub

Private Sub btnEstimateUpdate_Click()
    UpdateEstimate
End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnOrderListDelete_Click()
    DeleteOrderList
End Sub

Private Sub Frame4_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub

Private Sub frmOrder_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub

Private Sub UserForm_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub

Private Sub lswCustomerAutoComplete_DblClick()
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� �Ŵ����� �̵�
    With Me.lswCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
        End If
    End With
End Sub

Private Sub lswManagerAutoComplete_DblClick()
    '����ڸ� ���� �־��ְ� ��Ŀ���� ���������� �̵�
    With Me.lswManagerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtManager.value = .selectedItem.Text
            .Visible = False
            Me.txtEstimateName.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderList_DblClick()

    Dim i As Integer
    Dim pos As Integer
    
    With Me.lswOrderList
        headerIndex = 0
        For i = 1 To .ColumnHeaders.count
            pos = .ColumnHeaders(i).Left
            If mouseX < pos Then
                headerIndex = i - 1
                Exit For
            End If
        Next
        
        If headerIndex = 0 Then
            If Not .selectedItem Is Nothing Then
                clickOrderId = .selectedItem.Text
                
                If isFormLoaded("frmOrderUpdate") = True Then
                    Unload frmOrderUpdate
                End If
                frmOrderUpdate.Show (False)
            End If
        ElseIf headerIndex = 12 Then
            '�ݾ��� ������ �� ����
        Else
            ' ���� ������ ���� �����س���
            If Not beforeSelectedItem Is Nothing Then
                Set beforeSelectedItem = Nothing
            End If
            Set beforeSelectedItem = .selectedItem
            
            SelectOrderListColumn
        End If
    End With

End Sub

Private Sub lswPaymentList_DblClick()

    Dim i As Integer
    Dim pos As Integer
    
    With Me.lswPaymentList
        headerIndex = 0
        For i = 1 To .ColumnHeaders.count
            pos = .ColumnHeaders(i).Left
            If mouseX < pos Then
                headerIndex = i - 1
                Exit For
            End If
        Next
        
        If headerIndex = 12 Then
            '�ΰ����� ������ �� ����
        ElseIf headerIndex >= 4 Then
            ' ���� ������ ���� �����س���
            If Not beforeSelectedItem Is Nothing Then
                Set beforeSelectedItem = Nothing
            End If
            Set beforeSelectedItem = .selectedItem
            
            SelectPaymentListColumn
        End If
    End With

End Sub

Private Sub txtManager_Enter()
    '�ڵ��ϼ� ����Ʈ���� ���ؼ� �Ѿ���� ���
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtEstimateName_Enter()
    '�ڵ��ϼ� ����Ʈ���� ���ؼ� �Ѿ���� ���
    With Me.lswManagerAutoComplete
        If .Visible = True Then
            Me.txtManager.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub lbl2AcceptedDate_Enter()
    Me.txtAcceptedDate.SetFocus
End Sub

Private Sub lbl2BidDate_Enter()
    Me.txtBidDate.SetFocus
End Sub

Private Sub lbl2DeliveryDate_Enter()
    Me.txtDeliveryDate.SetFocus
End Sub

Private Sub lbl2EstimateDate_Enter()
    Me.txtEstimateDate.SetFocus
End Sub

Private Sub UserForm_Activate()
    If bInitialIzed = False Then
        Me.txtSize.SetFocus
    End If
    bInitialIzed = True
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = vbKeyReturn Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtManager.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸� ���� �Է�ĭ���� �̵�
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtManager.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = vbKeyDown Then
            '�Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
End Sub

Private Sub txtCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '�ŷ�ó �ڵ��ϼ� ó��
    With Me.lswCustomerAutoComplete
        If Me.txtCustomer.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '�����ŷ�ó DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtEstimateCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.value, 1, False)
            If isEmpty(db) Then
                .Visible = False
            Else
                For i = 1 To UBound(db)
                    .ListItems.Add , , db(i, 1)
                    If i = 8 Then Exit For
                Next
            End If
            
        End If
    End With
End Sub

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswManagerAutoComplete
        If KeyCode = vbKeyReturn Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtEstimateName.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸� ���� �Է�ĭ���� �̵�
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtEstimateName.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = vbKeyDown Then
            '�Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
    
End Sub

Private Sub txtSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub chkDividePay_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.btnEstimateUpdate.SetFocus
    End If
End Sub


Private Sub txtManager_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '����� �ڵ��ϼ� ó��
    With Me.lswManagerAutoComplete
        If Me.txtManager.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '��������� DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtEstimateManager, True)
            db = Filtered_DB(db, Me.txtManager.value, 1, False)
            If isEmpty(db) Then
                .Visible = False
            Else
                For i = 1 To UBound(db)
                    .ListItems.Add , , db(i, 1)
                    If i = 8 Then Exit For
                Next
            End If
            
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� �Ŵ����� �̵�
    If KeyCode = vbKeyReturn Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
End Sub

Private Sub lswManagerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '����� ���� �� ����Ű ������ �� ���� ����ڸ� �־��ְ� ��Ŀ���� ����(������)���� �̵�
    If KeyCode = vbKeyReturn Then
        With Me.lswManagerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtManager.value = .selectedItem.Text
                .Visible = False
                Me.txtEstimateName.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub btnProduction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub


Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyTab Then
'        MsgBox Me.txtMemo.CurTargetX & ", " & Me.txtMemo.CurX
'        KeyCode = 0
'    End If
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cboOrderCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswOrderList
        If KeyCode = vbKeyReturn Then
            If headerIndex = 0 Then headerIndex = 4
            OrderListUpdate headerIndex
            Me.cboOrderCategory.Visible = False
            Me.frmEdit.Visible = False
            .SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If headerIndex = 0 Then headerIndex = 4
            OrderListUpdate headerIndex
            headerIndex = headerIndex + 1
            SelectOrderListColumn
            KeyCode = 0
        End If
    End With
End Sub

Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long
    
    With Me.lswOrderList
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
            '���氪�� DB�� ȭ�鿡 �ݿ�
            OrderListUpdate headerIndex
            
            '����Ű - ���� �ٲ���. ����ĭ���� �̵����� ����
            If KeyCode = vbKeyReturn Then
                Me.txtEdit.Visible = False
                Me.frmEdit.Visible = False
                .SetFocus
            ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyRight Then
                '��Ű, ������ ȭ��ǥŰ
                If headerIndex = 18 Then
                    Me.txtEdit.Visible = False
                    Me.frmEdit.Visible = False
                    .SetFocus
                ElseIf headerIndex = 11 Then
                    headerIndex = headerIndex + 2
                    SelectOrderListColumn
                    KeyCode = 0
                Else
                    headerIndex = headerIndex + 1
                    SelectOrderListColumn
                    KeyCode = 0
                End If
            ElseIf KeyCode = vbKeyUp Then
                '����ȭ��ǥŰ
                '����Ʈ �� ó���� �ƴϸ� ��ĭ���� �̵�
                With Me.lswOrderList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = 1 Then
                                Me.txtEdit.Visible = False
                                Me.frmEdit.Visible = False
                                .SetFocus
                            Else
                                .ListItems(i).Selected = False
                                .ListItems(i - 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectOrderListColumn
                                KeyCode = 0
                                Exit For
                            End If
                        End If
                    Next
                End With
            ElseIf KeyCode = vbKeyDown Then
                '�Ʒ�ȭ��ǥŰ
                With Me.lswOrderList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = .ListItems.count Then
                                '�� �������̸� ������
                                Me.txtEdit.Visible = False
                                Me.frmEdit.Visible = False
                                .SetFocus
                                Exit For
                            Else
                                '����Ʈ �� �������� �ƴϸ� ��ĭ �Ʒ��� �̵�
                                .ListItems(i).Selected = False
                                .ListItems(i + 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectOrderListColumn
                                Exit For
                            End If
                        End If
                    Next
                End With
                KeyCode = 0
            ElseIf KeyCode = vbKeyLeft Then
                '����ȭ��ǥŰ
                '�� ó���� �ƴϸ� ��ĭ �������� �̵�
                If headerIndex <= 4 Then
                    Me.txtEdit.Visible = False
                    Me.frmEdit.Visible = False
                    .SetFocus
                Else
                    If headerIndex = 13 Then
                        headerIndex = headerIndex - 2   '�ݾ� �ʵ� �ǳʶٱ� ���ؼ� -2 ����
                    Else
                        headerIndex = headerIndex - 1
                    End If
                    SelectOrderListColumn
                    KeyCode = 0
                End If
            End If
        
        ElseIf KeyCode = vbKeyEscape Then
            'ESCŰ
            Me.txtEdit.Visible = False
            Me.frmEdit.Visible = False
        End If
    End With
End Sub

Sub OrderListUpdate(headerIndex)
    Dim orderPrice As Long
    
    With Me.lswOrderList
        If .selectedItem Is Nothing Then
            Exit Sub
        End If
        
        If headerIndex = 4 Then
            If Me.cboOrderCategory.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.cboOrderCategory.value
                'DB ���̺� ����
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.cboOrderCategory.value
            End If
        Else
            If Me.txtEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '�Է°� ���� ����
                ConvertOrderListFormat Me.txtEdit, headerIndex
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtEdit.value
                'DB ���̺� ����
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.txtEdit.value
                
                '����,�ܰ� ������ ��쿡�� �ݾ׵� �����ؾ� ��
                If headerIndex = 9 Or headerIndex = 11 Then
                    orderPrice = CalculateOrderListPrice(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(orderPrice, "#,##0")
                    UpdateOrderListValue .selectedItem.Text, 12, orderPrice
                End If
                '���డ �Ѿ� ���
                Me.txtExecutionCost = Format(CalculateOrderListTotalCost, "#,##0")
                CalculateAcceptedMargin
            End If
        End If
    End With
End Sub

Sub PaymentListUpdate(headerIndex)
    Dim VAT As Variant
    
    With Me.lswPaymentList
        If .selectedItem Is Nothing Then
            Exit Sub
        End If
        
        If headerIndex = 10 Then
            If Me.cboEstimatePayMethod.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.cboEstimatePayMethod.value
                'DB ���̺� ����
                UpdatePaymentListValue .selectedItem.Text, headerIndex, Me.cboEstimatePayMethod.value
            End If
        Else
            If Me.txtPaymentEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '�Է°� ���� ����
                ConvertPaymentListFormat Me.txtPaymentEdit, headerIndex
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtPaymentEdit.value
                'DB ���̺� ����
                UpdatePaymentListValue .selectedItem.Text, headerIndex, Me.txtPaymentEdit.value
                
                If headerIndex = 7 Then
                    '�������� ��쿡 ȭ�鿡 '10��' �̷������� ������
                    .selectedItem.ListSubItems(6).Text = Format(Me.txtPaymentEdit.value, "mm" & "��")
                    .selectedItem.ListSubItems(12).Text = Me.txtPaymentEdit.value
                ElseIf headerIndex = 5 Or headerIndex = 8 Or headerIndex = 9 Then
                    '�Աݾ� ������ ��쿡�� �ΰ����� �����ؾ� ��
                    VAT = CalculatePaymentListVAT(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(VAT, "#,##0")
                    UpdatePaymentListValue .selectedItem.Text, 12, VAT
                End If
                '�հ�
                Me.txtPaid.value = Format(CalculatePaymentListTotalCost, "#,##0")
                CalculatePayment
            End If
        End If
    End With
End Sub


Private Sub cboEstimatePayMethod_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswPaymentList
        If KeyCode = vbKeyReturn Then
            If headerIndex = 0 Then headerIndex = 10
            PaymentListUpdate headerIndex
            Me.cboEstimatePayMethod.Visible = False
            Me.frmPaymentEdit.Visible = False
            .SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If headerIndex = 0 Then headerIndex = 10
            PaymentListUpdate headerIndex
            headerIndex = headerIndex + 1
            SelectPaymentListColumn
            KeyCode = 0
        End If
    End With
End Sub

Private Sub txtPaymentEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long
    
    With Me.lswPaymentList
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
            '���氪�� DB�� ȭ�鿡 �ݿ�
            PaymentListUpdate headerIndex

            '����Ű - ���� �ٲ���. ����ĭ���� �̵����� ����
            If KeyCode = vbKeyReturn Then
                Me.txtPaymentEdit.Visible = False
                Me.frmPaymentEdit.Visible = False

                Me.lswPaymentList.SetFocus
            ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyRight Then
                '��Ű, ������ ȭ��ǥŰ
                If headerIndex = 11 Then
                    Me.txtPaymentEdit.Visible = False
                    Me.frmPaymentEdit.Visible = False
                    Me.lswPaymentList.SetFocus
                Else
                    headerIndex = headerIndex + 1
                    SelectPaymentListColumn
                    KeyCode = 0
                End If
            ElseIf KeyCode = vbKeyUp Then
                '����ȭ��ǥŰ
                '����Ʈ �� ó���� �ƴϸ� ��ĭ���� �̵�
                With Me.lswPaymentList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = 1 Then
                                Me.txtPaymentEdit.Visible = False
                                Me.frmPaymentEdit.Visible = False
                                Me.lswPaymentList.SetFocus
                            Else
                                .ListItems(i).Selected = False
                                .ListItems(i - 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectPaymentListColumn
                                KeyCode = 0
                                Exit For
                            End If
                        End If
                    Next
                End With
            ElseIf KeyCode = vbKeyDown Then
                '�Ʒ�ȭ��ǥŰ
                With Me.lswPaymentList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = .ListItems.count Then
                                '�� �������̸� ������
                                Me.txtPaymentEdit.Visible = False
                                Me.frmPaymentEdit.Visible = False
                                Me.lswPaymentList.SetFocus
                                Exit For
                            Else
                                '����Ʈ �� �������� �ƴϸ� ��ĭ �Ʒ��� �̵�
                                .ListItems(i).Selected = False
                                .ListItems(i + 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectPaymentListColumn
                                Exit For
                            End If
                        End If
                    Next
                End With
                KeyCode = 0
            ElseIf KeyCode = vbKeyLeft Then
                '����ȭ��ǥŰ
                '�� ó���� �ƴϸ� ��ĭ �������� �̵�
                If headerIndex <= 4 Then
                    Me.txtPaymentEdit.Visible = False
                    Me.frmPaymentEdit.Visible = False
                    Me.lswPaymentList.SetFocus
                Else
                    headerIndex = headerIndex - 1
                    SelectPaymentListColumn
                    KeyCode = 0
                End If
            End If
        
        ElseIf KeyCode = vbKeyEscape Then
            'ESCŰ
            Me.txtPaymentEdit.Visible = False
            Me.frmPaymentEdit.Visible = False
        End If
    End With
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    '�ؽ�Ʈ�ڽ��� ��Ŀ���� �ξ�� AfterUpdate�� ����
    Me.txtEstimateDate.SetFocus
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtBidDate.SetFocus
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtInsuranceDate.SetFocus
    GetCalendarDate Me.txtInsuranceDate
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtAcceptedDate.SetFocus
    GetCalendarDate Me.txtAcceptedDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtDueDate.SetFocus
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtDeliveryDate.SetFocus
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxinvoiceDate
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "��")
End Sub


Private Sub txtManagementID_AfterUpdate()
    Dim li As ListItem
    
    '�ʱ�ȭ �ÿ��� DB�� �����ϴ� �� ����
    If bInitialIzed = False Then Exit Sub
    
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
    
    '�������� DB �о����
    db = Get_DB(shtEstimate)
    
    '������ ������ȣ�� �ִ��� üũ
    blnUnique = IsUnique(db, Me.txtManagementID.value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    '������ ������ ���� DB�� ����
    UpdateEstimateOrderValue "������ȣ", Me.txtManagementID.value
    
    '������ȣ�� ���ֳ��� �ǵ鵵 DB���� ����
    For Each li In lswOrderList.ListItems
    
    Next
    
    '������ȣ ������ �Ǵ� ��� ����Ͽ� �ٲ���
    orgManagementID = Me.txtManagementID.value
    
End Sub

Private Sub txtEstimateName_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtEstimateName.value = Trim(Me.txtEstimateName.value)
    
    UpdateEstimateOrderValue "������", Me.txtEstimateName.value
End Sub

Private Sub txtMemo_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtMemo.value = Trim(Me.txtMemo.value)
    
    UpdateEstimateOrderValue "�޸�", Me.txtMemo.value
End Sub

Private Sub txtAmount_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    If Me.txtAmount.value <> "" Then
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '���� 1,000�ڸ� �ĸ� ó��
            Me.txtAmount.value = Format(Me.txtAmount.value, "#,##0")
        End If
    End If
    
    '��� �ʵ� ���
    CalculateEstimatePrice
    
    'DB ����
    UpdateEstimateOrderValue "����", Me.txtAmount.value
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    If Me.txtUnitPrice.value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�ܰ� 1,000�ڸ� �ĸ� ó��
            Me.txtUnitPrice.value = Format(Me.txtUnitPrice.value, "#,##0")
        End If
    End If
    
    '��� �ʵ� ���
    CalculateEstimatePrice
    
    'DB ����
    '�ܰ��� ����DB���� ����
    UpdateEstimateValue "�ܰ�", Me.txtUnitPrice.value
End Sub

Private Sub txtEstimatePrice_Change()
    If bInitialIzed = False Then Exit Sub
    
    '�����ݾ��� ����DB�� ������
     UpdateEstimateValue "�ݾ�", Me.txtEstimatePrice.value
End Sub

Private Sub txtBidPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     If Me.txtBidPrice.value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.value) Then
            Me.txtBidPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�����ݾ� 1,000�ڸ� �ĸ� ó��
            Me.txtBidPrice.value = Format(Me.txtBidPrice.value, "#,##0")
        End If
    End If
    
    '��� �ʵ� ���
    CalculateBidMargin
    
    UpdateEstimateValue "�����ݾ�", Me.txtBidPrice.value
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     If Me.txtAcceptedPrice.value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtAcceptedPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            Me.txtAcceptedPrice.value = Format(Me.txtAcceptedPrice.value, "#,##0")
            
            CalculateAcceptedMargin
            CalculatePayment
            
            UpdateEstimateValue "���ֱݾ�", Me.txtAcceptedPrice.value
            UpdateOrderValue "�ݾ�", Me.txtAcceptedPrice.value
            If IsNumeric(Me.txtAmount.value) Then
                UpdateOrderValue "�ܰ�", CLng(Me.txtAcceptedPrice.value) / CLng(Me.txtAmount.value)
            End If
        End If
    End If
End Sub

Sub UpdateProductionTotalCost(fieldValue)
    Me.txtProductionTotalCost.value = fieldValue
End Sub

'Enable=False �� �ؽ�Ʈ�ڽ��� AfterUpdate�� ������ �ȵǴ� ��� ����. �̺�Ʈ�� Change�� �ٲ�
Private Sub txtProductionTotalCost_Change()
    If bInitialIzed = False Then Exit Sub
    
     '��� �ʵ� ���
    CalculateBidMargin
            
    'DB �ݿ�
    UpdateEstimateOrderValue "���డ(����)", Me.txtProductionTotalCost.value
End Sub

'Enable=False �� �ؽ�Ʈ�ڽ��� AfterUpdate�� ������ �ȵǴ� ��� ����. �̺�Ʈ�� Change�� �ٲ�
Private Sub txtExecutionCost_Change()
    If bInitialIzed = False Then Exit Sub
    
     CalculateAcceptedMargin
    
    UpdateEstimateOrderValue "���డ", Me.txtExecutionCost.value
End Sub

Private Sub txtPaid_Change()
    If bInitialIzed = False Then Exit Sub
    
    CalculatePayment
    
    UpdateEstimateOrderValue "�Աݾ�", Me.txtPaid.value
End Sub

Private Sub txtRemaining_Change()
    If bInitialIzed = False Then Exit Sub
    
    UpdateEstimateOrderValue "���Աݾ�", Me.txtRemaining.value
End Sub

Private Sub txtVAT_Change()
    If bInitialIzed = False Then Exit Sub
    
    UpdateEstimateOrderValue "�ΰ���", Me.txtVAT.value
End Sub

Private Sub chkVAT_AfterUpdate()
    If bInitialIzed = False Then Exit Sub

    If Me.chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        CalculatePayment
    End If
    
    UpdateEstimateOrderValue "�ΰ�������", Me.chkVAT.value
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtAcceptedDate.value = ConvertDateFormat(Me.txtAcceptedDate.value)
    
    UpdateEstimateOrderValue "����", Me.txtAcceptedDate.value
End Sub

Private Sub cboCategory_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "�з�1", Me.cboCategory.value
End Sub

Private Sub txtAcceptedMargin_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "��������", Me.txtAcceptedMargin.value
End Sub

Private Sub txtAcceptedMarginRate_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "������", Me.txtAcceptedMarginRate.value
End Sub

Private Sub txtBidMargin_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "����(����)", Me.txtBidMargin.value
End Sub

Private Sub txtBidMarginRate_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "������(����)", Me.txtBidMarginRate.value
End Sub

Private Sub txtLinkedID_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "������ȣ", Me.txtLinkedID.value
End Sub
Private Sub cboUnit_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.cboUnit.value = Trim(Me.cboUnit.value)
    
    UpdateEstimateOrderValue "����", Me.cboUnit.value
End Sub

Private Sub txtBidDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtBidDate.value = ConvertDateFormat(Me.txtBidDate.value)
    
    UpdateEstimateOrderValue "����", Me.txtBidDate.value
End Sub

Private Sub txtCustomer_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
    
    UpdateEstimateOrderValue "�ŷ�ó", Me.txtCustomer.value
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtDeliveryDate.value = ConvertDateFormat(Me.txtDeliveryDate.value)
    
    UpdateEstimateOrderValue "��ǰ", Me.txtDeliveryDate.value
End Sub

Private Sub txtDueDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtDueDate.value = ConvertDateFormat(Me.txtDueDate.value)
    
    UpdateEstimateOrderValue "����", Me.txtDueDate.value
End Sub

Private Sub txtEstimateDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtEstimateDate.value = ConvertDateFormat(Me.txtEstimateDate.value)
    
    UpdateEstimateOrderValue "����", Me.txtEstimateDate.value
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtInsuranceDate.value = ConvertDateFormat(Me.txtInsuranceDate.value)
    
    UpdateEstimateOrderValue "����", Me.txtInsuranceDate.value
End Sub

Private Sub txtManager_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtManager.value = Trim(Me.txtManager.value)
    
    UpdateEstimateOrderValue "�����", Me.txtManager.value
End Sub

Private Sub txtSize_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtSize.value = Trim(Me.txtSize.value)
    
    UpdateEstimateOrderValue "�԰�", Me.txtSize.value
End Sub



Private Sub txtTaxinvoiceDate_AfterUpdate()
    Me.txtTaxinvoiceDate.value = Trim(Me.txtTaxinvoiceDate.value)
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = Trim(Me.txtPaymentDate.value)
End Sub

Private Sub txtExpectPaymentMonth_AfterUpdate()
    Dim pos As Long
    Dim M As String

    Me.txtExpectPaymentMonth.value = Trim(Me.txtExpectPaymentMonth.value)

    If Me.txtExpectPaymentMonth = "" Then Exit Sub
    
    pos = InStr(Me.txtExpectPaymentMonth, "��")
    If pos <> 0 Then
        M = Left(Me.txtExpectPaymentMonth, pos - 1)
        Me.txtExpectPaymentDate.value = DateSerial(Year(Date), M, 1)
        Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "��")
        Exit Sub
    End If
    
    If IsNumeric(Me.txtExpectPaymentMonth) Then
        Me.txtExpectPaymentDate.value = DateSerial(Year(Date), Me.txtExpectPaymentMonth, 1)
        Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "��")
        Exit Sub
    End If
    
    Me.txtExpectPaymentDate.value = Me.txtExpectPaymentMonth
    Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "��")
     
End Sub

Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = Trim(Me.txtSpecificationDate.value)
End Sub


Private Sub txtEdit_AfterUpdate()
    '��Ű�� ����Ű�� �ƴ� ���콺�� Ŭ���ؼ� ����� ���: currentEditText�� �����
    If headerIndex > 4 And headerIndex < Me.lswOrderList.ColumnHeaders.count Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.txtEdit.value <> currentEditText Then
                OrderListUpdate headerIndex
                headerIndex = 0
                currentEditText = ""
            End If
        End If
    End If
    
End Sub

Private Sub cboOrderCategory_AfterUpdate()
    If headerIndex = 4 Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.cboOrderCategory.value <> currentCboText Then
                OrderListUpdate headerIndex
                headerIndex = 0
                currentCboText = ""
            End If
        End If
    End If
End Sub

Private Sub txtPaymentEdit_AfterUpdate()
    '��Ű�� ����Ű�� �ƴ� ���콺�� Ŭ���ؼ� ����� ���: currentEditText�� ����ؾ� ��
    If headerIndex >= 4 And headerIndex < Me.lswPaymentList.ColumnHeaders.count Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.txtPaymentEdit.value <> currentEditText Then
                PaymentListUpdate headerIndex
                headerIndex = 0
                currentEditText = ""
            End If
        End If
    End If
End Sub

Private Sub cboEstimatePayMethod_AfterUpdate()
    If headerIndex = 10 Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.cboEstimatePayMethod.value <> currentCboText Then
                PaymentListUpdate headerIndex
                headerIndex = 0
                currentCboText = ""
            End If
        End If
    End If
End Sub

Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub


Private Sub UserForm_Terminate()
    bInitialIzed = False
End Sub
