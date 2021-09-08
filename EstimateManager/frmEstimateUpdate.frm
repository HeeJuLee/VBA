VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "���� ����"
   ClientHeight    =   12240
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


Private Sub frmOrder_Click()

End Sub

Private Sub UserForm_Activate()
    Me.txtManagementID.SetFocus
End Sub

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
        
        currentEstimateId = shtEstimateAdmin.Cells(cRow, 2)
    End If
    
     '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            If contr.Name Like "lbl2*" Then
            Else
                contr.BackColor = RGB(242, 242, 242)
            End If
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

    Me.txtID.value = estimate(1)    'ID
    Me.txtEstimateName.value = estimate(6)  '������
    Me.txtManagementID.value = estimate(2)    '������ȣ
    Me.txtLinkedID.value = estimate(3)  '�����ȣ
    
    Me.txtCustomer = estimate(4)   '�ŷ�ó
    Me.txtManager = estimate(5)   '�����
    
    Me.txtSize.value = estimate(7)  '�԰�
    Me.txtAmount.value = Format(estimate(8), "#,##0")   '����
    InitializeCboUnit
    Me.cboUnit.value = Trim(estimate(9))  '����, ID�� �����Ƿ� ���� value ������ ���õ�
    Me.txtUnitPrice.value = Format(estimate(10), "#,##0")     '�����ܰ�
    Me.txtEstimatePrice.value = Format(estimate(11), "#,##0")     '�����ݾ�
    
    Me.txtEstimateDate.value = estimate(12)    '��������
    Me.txtBidDate.value = estimate(13)    '��������
    Me.txtAcceptedDate.value = estimate(14)    '��������
    Me.txtDeliveryDate.value = estimate(15)    '��ǰ����
    Me.txtInsuranceDate.value = estimate(16)    '��������
    
    Me.txtProductionTotalCost.value = Format(estimate(17), "#,##0")   '���డ
    Me.txtBidPrice.value = Format(estimate(18), "#,##0")    '������
    Me.txtBidMargin.value = Format(estimate(19), "#,##0")    '����
    Me.txtBidMarginRate.value = Format(estimate(20), "0.0%")    '������
    Me.txtAcceptedPrice.value = Format(estimate(21), "#,##0")    '���ֱݾ�
    Me.txtAcceptedMargin.value = Format(estimate(22), "#,##0")   '��������
    
    Me.txtInsertDate.value = estimate(23)    '�������
    Me.txtUpdateDate.value = estimate(24)    '��������
    
    InitializeCboCategory
    Me.cboCategory.value = Trim(estimate(25))   '�з�
    Me.txtDueDate.value = estimate(26)              '������
    Me.txtSpecificationDate.value = estimate(27)    '�ŷ�����
    Me.txtTaxinvoiceDate.value = estimate(28)    '���ݰ�꼭
    Me.txtPaymentDate.value = estimate(29)    '��������
    Me.txtExpectPaymentDate.value = estimate(30)  '���������
    Me.txtExpectPaymentMonth.value = Format(estimate(30), "mm" & "��")  '���������
    Me.txtVAT.value = Format(estimate(31), "#,##0")    '�ΰ���
    Me.txtMemo.value = Trim(estimate(32))     '�����޸�
    Me.chkVAT.value = estimate(33)      '�ΰ��� ���� ����
    
    Me.txtPaid.value = Format(estimate(34), "#,##0")      '�Աݾ�
    Me.txtRemaining.value = Format(estimate(35), "#,##0")      '���Աݾ�
    Me.chkDividePay.value = estimate(36)      '���Ұ��� ����
    If chkDividePay.value = True Then
        Me.btnPayment.Enabled = True
    Else
        Me.btnPayment.Enabled = False
    End If
    
    '�������� ��������xls�� ����xls���� �޴��� �ٸ��� ����
    '������ �����޸�� ���ָ޸�� ���� ����� ��.
    '�����޸�� ���ָ޸� �ٸ� ���� ���� �����
    '(�����޸� = �����޸� + ���ָ޸�) �̷��� ���߰� ���� �� �����ʿ��� �����ϰ� �޸� ���� ����
    acceptedMemo = Trim(estimate(37))
    If Me.txtMemo.value <> acceptedMemo Then
        If Me.txtMemo.value = "" Then
            Me.txtMemo.value = acceptedMemo
        Else
            Me.txtMemo.value = Me.txtMemo.value & vbCrLf & acceptedMemo
        End If
    End If
    
    '���� ID (ID_����)
    Me.txtAcceptedID.value = estimate(38)
    If Me.txtAcceptedID.value = "" Then
        '����ID�� ������ ���ְ��� ��Ʈ�� unable ��Ŵ
        frmOrder.Visible = False
        btnAcceptedInsert.Visible = True
        frmEstimateUpdate.Height = 280
    Else
        frmOrder.Visible = True
        btnAcceptedInsert.Visible = False
    End If
    
    '���� �� ������ȣ
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '���� ��Ȳ
    InitializeLswCustomerAutoComplete   '�ŷ�ó �ڵ��ϼ�
    InitializeLswManagerAutoComplete    '����� �ڵ��ϼ�
    
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
    If Not IsEmpty(db) Then
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

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� ���� ������ �о��
    db = Get_DB(shtOrder)
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.value, 28, True)
    End If
    If Not IsEmpty(db) Then
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
        .ColumnHeaders.Add , , "������", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 30
        
        '.ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not IsEmpty(db) Then
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
                li.ListSubItems.Add , , db(i, 22)       '������
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

Sub UpdateEstimate()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckEstimateUpdateValidation = False Then
        Exit Sub
    End If

    '�������� DB �о����
    db = Get_DB(shtEstimate)
    
    '������ ������ȣ�� �ִ��� üũ
    blnUnique = IsUnique(db, Me.txtManagementID.value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    '���� ���̺� ������Ʈ
    Update_Record shtEstimate, Me.txtID.value, _
        Me.txtManagementID.value, Me.txtLinkedID.value, _
        Me.txtCustomer.value, Me.txtManager.value, _
        Me.txtEstimateName.value, Me.txtSize.value, _
        Me.txtAmount.value, Me.cboUnit.value, _
        Me.txtUnitPrice.value, Me.txtEstimatePrice.value, _
        Me.txtEstimateDate.value, Me.txtBidDate.value, _
        Me.txtAcceptedDate.value, Me.txtDeliveryDate.value, _
        Me.txtInsuranceDate.value, Me.txtProductionTotalCost.value, _
        Me.txtBidPrice.value, Me.txtBidMargin.value, _
        Me.txtBidMarginRate.value, Me.txtAcceptedPrice.value, _
        Me.txtAcceptedMargin.value, _
        Me.txtInsertDate.value, Date, _
        Me.cboCategory.value, Me.txtDueDate.value, _
        Me.txtSpecificationDate.value, Me.txtTaxinvoiceDate.value, Me.txtPaymentDate.value, Me.txtExpectPaymentDate.value, _
        Me.txtVAT.value, Me.txtMemo.value, Me.chkVAT.value, _
        Me.txtPaid.value, Me.txtRemaining.value, Me.chkDividePay, ""
    
    '���� ���̺� ������Ʈ
    If Me.txtAcceptedID.value <> "" Then
        Update_Record shtOrder, Me.txtAcceptedID.value, _
        , Me.cboCategory.value, , _
        Me.txtManagementID.value, Me.txtCustomer.value, _
        Me.txtEstimateName.value, Me.txtManager.value, _
        Me.txtSize.value, Me.txtAmount.value, _
        Me.cboUnit.value, Me.txtUnitPrice, _
        Me.txtEstimatePrice.value, , _
        Me.txtAcceptedDate.value, , Me.txtDueDate.value, _
        , Me.txtDeliveryDate.value, _
        Me.txtSpecificationDate.value, Me.txtTaxinvoiceDate.value, Me.txtPaymentDate.value, Me.txtExpectPaymentDate.value, _
        , Me.txtVAT.value, _
        , Date, _
        , Me.txtMemo.value, Me.chkVAT.value
    End If
    
    '������ȣ ������ �Ǵ� ��� ����Ͽ� �ٲ���
    orgManagementID = Me.txtManagementID.value
    
    UpdateShtEstimate Me.txtID.value
    
    UpdateShtOrder Me.txtID.value
    
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
            Me.txtUnitPrice.value, _
            Me.txtEstimatePrice.value, _
            , _
            , , , , , _
            , , , , _
            , , _
            Date, , _
            CLng(Me.txtID.value), , False

    '����� ����ID�� ���� ���̺� ������Ʈ
    Update_Record_Column shtEstimate, Me.txtID, "ID_����", Get_LastID(shtOrder)
    
    '���� ���� ���
    Unload frmEstimateUpdate
    frmEstimateUpdate.Show (False)
    
End Sub

'���� ����Ʈ�� �� ����
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
            fieldName = "��������"
        Case 14  '����
            fieldName = "��������"
        Case 15  '�԰�
            fieldName = "�԰�����"
        Case 16  '����
            fieldName = "��������"
        Case 17  '��꼭
            fieldName = "��꼭����"
        Case 18  '������
            fieldName = "��������"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtOrder, id, fieldName, value
        Update_Record_Column shtOrder, id, "��������", Date
    End If

End Sub


Sub SelectOrderListColumn()
    Dim ItemSel    As ListItem
    
    If Not lswOrderList.selectedItem Is Nothing Then
        If headerIndex = lswOrderList.ColumnHeaders.count Then
            frmEdit.Visible = False
            txtEdit.Visible = False
        End If
        
        If headerIndex > 0 And headerIndex < lswOrderList.ColumnHeaders.count Then
        
            Set ItemSel = lswOrderList.selectedItem
        
            With frmEdit
                .Visible = True
                .top = ItemSel.top + lswOrderList.top
                .Left = lswOrderList.ColumnHeaders(headerIndex).Left + lswOrderList.Left
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width
                .Height = ItemSel.Height + 10
                .ZOrder msoBringToFront
            End With
            
            With Me.txtEdit
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width - 2
                .Height = lswOrderList.selectedItem.Height + 3
                .SelLength = Len(.Text)
            End With
        End If
    End If

End Sub

Sub DeleteOrderList()
    Dim li As ListItem
    Dim count As Long
    Dim YN As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "������ ���ָ� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    YN = MsgBox("������ " & count & "�� ���ָ� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If YN = vbNo Then Exit Sub

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

Sub BatchUpdateOrderdate()
    Dim li As ListItem
    Dim count As Long
    Dim YN As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "�ϰ� ������ ���ָ� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    If isFormLoaded("frmOrderDateUpdate") = True Then
        Unload frmOrderDateUpdate
    End If
    frmOrderDateUpdate.Show (False)
    
End Sub

Sub UpdateShtEstimate(estimateId)
    Dim findRow As Long
    
    findRow = isExistInSheet(shtEstimateAdmin.Range("B6"), estimateId)
    If findRow <> 0 Then
        shtEstimateAdmin.Cells(findRow, 4).value = Me.txtManagementID.value
        shtEstimateAdmin.Cells(findRow, 5).value = Me.txtCustomer.value
        shtEstimateAdmin.Cells(findRow, 6).value = Me.txtManager.value
        shtEstimateAdmin.Cells(findRow, 7).value = Me.cboCategory.value
        shtEstimateAdmin.Cells(findRow, 8).value = Me.txtEstimateName.value
        shtEstimateAdmin.Cells(findRow, 9).value = Me.txtSize.value
        shtEstimateAdmin.Cells(findRow, 10).value = Me.txtAmount.value
        shtEstimateAdmin.Cells(findRow, 11).value = Me.cboUnit.value
        shtEstimateAdmin.Cells(findRow, 12).value = Me.txtUnitPrice.value
        shtEstimateAdmin.Cells(findRow, 13).value = Me.txtEstimatePrice.value
        shtEstimateAdmin.Cells(findRow, 14).value = Me.txtEstimateDate.value
        shtEstimateAdmin.Cells(findRow, 15).value = Me.txtBidDate.value
        shtEstimateAdmin.Cells(findRow, 16).value = Me.txtAcceptedDate.value
        shtEstimateAdmin.Cells(findRow, 17).value = Me.txtDueDate.value
        shtEstimateAdmin.Cells(findRow, 18).value = Me.txtDeliveryDate.value
        shtEstimateAdmin.Cells(findRow, 19).value = Me.txtInsuranceDate.value
        shtEstimateAdmin.Cells(findRow, 20).value = Me.txtProductionTotalCost.value
        shtEstimateAdmin.Cells(findRow, 21).value = Me.txtBidPrice.value
        shtEstimateAdmin.Cells(findRow, 22).value = Me.txtBidMargin.value
        shtEstimateAdmin.Cells(findRow, 23).value = Me.txtBidMarginRate.value
        shtEstimateAdmin.Cells(findRow, 24).value = Me.txtAcceptedPrice.value
        shtEstimateAdmin.Cells(findRow, 25).value = Me.txtAcceptedMargin.value
        shtEstimateAdmin.Cells(findRow, 26).value = Me.txtSpecificationDate.value
        shtEstimateAdmin.Cells(findRow, 27).value = Me.txtTaxinvoiceDate.value
        shtEstimateAdmin.Cells(findRow, 28).value = Me.txtPaymentDate.value
        shtEstimateAdmin.Cells(findRow, 29).value = Me.txtExpectPaymentDate.value
        shtEstimateAdmin.Cells(findRow, 30).value = Me.txtVAT.value
        shtEstimateAdmin.Cells(findRow, 31).value = Me.txtInsertDate.value
        shtEstimateAdmin.Cells(findRow, 32).value = Date
    End If
End Sub

Sub UpdateShtOrder(orderId)
    Dim findRow As Long
    
    findRow = isExistInSheet(shtOrderAdmin.Range("C6"), orderId)
    If findRow <> 0 Then
        shtOrderAdmin.Cells(findRow, 5).value = Me.txtManagementID.value
        shtOrderAdmin.Cells(findRow, 6).value = Me.cboCategory.value
        shtOrderAdmin.Cells(findRow, 8).value = Me.txtCustomer.value
        shtOrderAdmin.Cells(findRow, 9).value = Me.txtEstimateName.value
        shtOrderAdmin.Cells(findRow, 10).value = Me.txtManager.value
        shtOrderAdmin.Cells(findRow, 11).value = Me.txtSize.value
        shtOrderAdmin.Cells(findRow, 12).value = Me.txtAmount.value
        shtOrderAdmin.Cells(findRow, 13).value = Me.cboUnit.value
        shtOrderAdmin.Cells(findRow, 14).value = Me.txtUnitPrice.value
        shtOrderAdmin.Cells(findRow, 15).value = Me.txtEstimatePrice.value
        shtOrderAdmin.Cells(findRow, 17).value = Me.txtAcceptedDate.value
        shtOrderAdmin.Cells(findRow, 19).value = Me.txtDueDate.value
        shtOrderAdmin.Cells(findRow, 21).value = Me.txtDeliveryDate.value
        shtOrderAdmin.Cells(findRow, 22).value = Me.txtSpecificationDate.value
        shtOrderAdmin.Cells(findRow, 23).value = Me.txtTaxinvoiceDate.value
        shtOrderAdmin.Cells(findRow, 24).value = Me.txtPaymentDate.value
        shtOrderAdmin.Cells(findRow, 25).value = Me.txtExpectPaymentDate.value
        shtOrderAdmin.Cells(findRow, 27).value = Me.txtVAT.value
        shtOrderAdmin.Cells(findRow, 28).value = Me.txtInsertDate.value
        shtOrderAdmin.Cells(findRow, 29).value = Date
    End If
End Sub

Sub UpdateShtOrderField(orderId, headerIndex, value)
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
            Case 18  '������
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


Sub CalculateEstimateUpdateCost()

    '�����ݾ� ���
    '�������� �����̸� �����ݾ��� �����ܰ�
    If Me.txtUnitPrice <> "" Then
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
        Me.txtBidMargin.value = 0
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

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.value, 2, True)
    
    'DB�� ���� ���� ���
    totalCost = 0
    If Not IsEmpty(db) Then
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
        Case 13, 14, 15, 16, 17, 18  '����, ����, �԰�, ����, ��꼭, ������ - ��¥ ��ȯ
            pos = InStr(value, "/")
            If pos > 0 Then
                M = Left(value, pos - 1)
                If Len(value) = pos Then
                    pos = 0
                Else
                    D = Mid(value, pos + 1)
                End If
            End If
            
            If pos > 0 Then
                textBox.Text = DateSerial(Year(Date), M, D)
            End If
    End Select
    
End Sub

Function isExistInSheet(startRng As Range, value) As Long
    Dim WS As Worksheet
    Dim lastRow As Long
    Dim col As Long
    Dim i As Long
    Set WS = startRng.Parent
    
    lastRow = startRng.End(xlDown).row
    col = startRng.Column
    
    If IsNumeric(value) Then
        value = CLng(value)
    End If
    
    isExistInSheet = 0
    For i = startRng.row To lastRow
        If WS.Cells(i, col) = value Then
            isExistInSheet = i
            Exit Function
        End If
    Next
End Function


Private Sub lswOrderList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
'    shtEstimateAdmin.Range("Q12").value = x
'    shtEstimateAdmin.Range("Q13").value = pointsPerPixelX * x
    mouseX = pointsPerPixelX * X
End Sub


Private Sub btnBatchUpdate_Click()
    BatchUpdateOrderdate
End Sub


Private Sub lswOrderList_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.value = ""
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

Private Sub btnProduction_Click()
    If isFormLoaded("frmProduction") Then
        Unload frmProduction
    End If
    frmProduction.Show (False)
End Sub


Private Sub btnAcceptedInsert_Click()
    InsertAccepted
End Sub

Private Sub btnPayment_Click()
    If isFormLoaded("frmPayment") Then
        Unload frmPayment
    End If
    frmPayment.Show (False)
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

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = 13 Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtManager.SetFocus
        ElseIf KeyCode = 9 Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸� ���� �Է�ĭ���� �̵�
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtManager.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = 40 Then
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
            If IsEmpty(db) Then
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
        If KeyCode = 13 Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtEstimateName.SetFocus
        ElseIf KeyCode = 9 Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸� ���� �Է�ĭ���� �̵�
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtEstimateName.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = 40 Then
            '�Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
    
End Sub

Private Sub chkDividePay_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
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
            If IsEmpty(db) Then
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
    If KeyCode = 13 Then
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
    If KeyCode = 13 Then
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
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub


'Private Sub btnPayHistoryInsert_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 9 Then
'        Me.btnEstimateUpdate.SetFocus
'    End If
'End Sub

Private Sub btnProduction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub

Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
    
    Me.txtMemo.value = Trim(Me.txtMemo.value)
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim orderPrice As Long
    Dim findRow As Long
    
    With Me.lswOrderList
        If KeyCode = 13 Or KeyCode = 9 Then
            If Me.txtEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '�Է°� ���� ����
                ConvertOrderListFormat Me.txtEdit, headerIndex
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtEdit.value
                'DB ���̺� ����
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.txtEdit.value
                '���� ���� ��Ʈ ����
                UpdateShtOrderField .selectedItem.Text, headerIndex, Me.txtEdit.value
    
                '����,�ܰ� ������ ��쿡�� �ݾ׵� �����ؾ� ��
                If headerIndex = 9 Or headerIndex = 11 Then
                    orderPrice = CalculateOrderListPrice(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(orderPrice, "#,##0")
                    UpdateOrderListValue .selectedItem.Text, 12, orderPrice
                    UpdateShtOrderField .selectedItem.Text, 12, orderPrice
                End If
                '���డ �Ѿ� ���
                Me.txtExecutionCost = Format(CalculateOrderListTotalCost, "#,##0")
                CalculateEstimateUpdateCost
            End If
            
            '����Ű - ���� �ٲ���. ����ĭ���� �̵����� ����
            '��Ű - �� �ٲ��ְ� ����ĭ�� txtEdit�� ������
            If KeyCode = 13 Then
                Me.txtEdit.Visible = False
                Me.frmEdit.Visible = False
                
                Me.lswOrderList.SetFocus
            Else
                If headerIndex = 11 Then
                    headerIndex = headerIndex + 2   '�ݾ� �ʵ� �ǳʶٱ� ���ؼ� +2 ����
                Else
                    headerIndex = headerIndex + 1
                End If
                SelectOrderListColumn
                
                '��Ŀ�� �ȳѾ���� ��
                KeyCode = 0
            End If
            
        ElseIf KeyCode = 40 Then
            '�Ʒ�ȭ��Ű
            KeyCode = 0
        ElseIf KeyCode = 27 Then
            'ESCŰ
            Me.txtEdit.Visible = False
            Me.frmEdit.Visible = False
        End If
    End With
End Sub


Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxinvoiceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "��")
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.txtEstimateName.value = Trim(Me.txtEstimateName.value)
End Sub

Private Sub txtAmount_AfterUpdate()

    If Me.txtAmount.value <> "" Then
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '���� 1,000�ڸ� �ĸ� ó��
            Me.txtAmount.value = Format(Me.txtAmount.value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�����ܰ� 1,000�ڸ� �ĸ� ó��
            Me.txtUnitPrice.value = Format(Me.txtUnitPrice.value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtBidPrice_AfterUpdate()
    
    If Me.txtBidPrice.value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.value) Then
            Me.txtBidPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�����ݾ� 1,000�ڸ� �ĸ� ó��
            Me.txtBidPrice.value = Format(Me.txtBidPrice.value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If Me.txtAcceptedPrice.value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtAcceptedPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            Me.txtAcceptedPrice.value = Format(Me.txtAcceptedPrice.value, "#,##0")
            
            CalculateEstimateUpdateCost
        End If
    End If
End Sub

Private Sub txtProductionTotalCost_AfterUpdate()
    
    If Me.txtProductionTotalCost.value <> "" Then
        If Not IsNumeric(Me.txtProductionTotalCost.value) Then
            Me.txtProductionTotalCost.value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '������డ 1,000�ڸ� �ĸ� ó��
            Me.txtProductionTotalCost.value = Format(Me.txtProductionTotalCost.value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtExecutionCost_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    Me.txtAcceptedDate.value = Trim(Me.txtAcceptedDate.value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
    Me.txtTaxinvoiceDate.value = Trim(Me.txtTaxinvoiceDate.value)
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = Trim(Me.txtPaymentDate.value)
    CalculateEstimateUpdateCost
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


Private Sub cboUnit_AfterUpdate()
    Me.cboUnit.value = Trim(Me.cboUnit.value)
End Sub

Private Sub txtBidDate_AfterUpdate()
    Me.txtBidDate.value = Trim(Me.txtBidDate.value)
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    Me.txtDeliveryDate.value = Trim(Me.txtDeliveryDate.value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.value = Trim(Me.txtDueDate.value)
End Sub

Private Sub txtEstimateDate_AfterUpdate()
    Me.txtEstimateDate.value = Trim(Me.txtEstimateDate.value)
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    Me.txtInsuranceDate.value = Trim(Me.txtInsuranceDate.value)
End Sub

Private Sub txtManager_AfterUpdate()
    Me.txtManager.value = Trim(Me.txtManager.value)
End Sub


Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub


Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = Trim(Me.txtSpecificationDate.value)
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub


Private Sub txtEdit_AfterUpdate()
    Dim orderPrice As Long
    Dim findRow As Long
    
    If headerIndex > 0 And headerIndex < Me.lswOrderList.ColumnHeaders.count Then
        '��Ű�� ����Ű�� �ƴ� ���콺�� Ŭ���ؼ� ����� ���: beforeSelectedItem�� ����ؾ� ��
        
        Debug.Print "AfterUpdate - headerIndex: " & headerIndex
        Debug.Print "AfterUpdate - ������: " & beforeSelectedItem.ListSubItems(headerIndex - 1)
        Debug.Print "AfterUpdate - ���氪: " & Me.txtEdit.value
        
        If Me.txtEdit.value <> beforeSelectedItem.ListSubItems(headerIndex - 1).Text Then
            '�Է°� ���� ����
            ConvertOrderListFormat Me.txtEdit, headerIndex
            beforeSelectedItem.ListSubItems(headerIndex - 1).Text = Me.txtEdit.value
            UpdateOrderListValue beforeSelectedItem.Text, headerIndex, Me.txtEdit.value
            UpdateShtOrderField beforeSelectedItem.Text, headerIndex, Me.txtEdit.value
                        
            '����,�ܰ� ������ ��쿡�� �ݾ׵� �����ؾ� ��
            If headerIndex = 9 Or headerIndex = 11 Then
                orderPrice = CalculateOrderListPrice(beforeSelectedItem)
                beforeSelectedItem.ListSubItems(11).Text = Format(orderPrice, "#,##0")
                UpdateOrderListValue beforeSelectedItem.Text, 12, orderPrice
                UpdateShtOrderField beforeSelectedItem.Text, 12, orderPrice
            End If
                
            '���డ �Ѿ� ���
            Me.txtExecutionCost = Format(CalculateOrderListTotalCost, "#,##0")
            CalculateEstimateUpdateCost
                
            headerIndex = 0
        End If
    End If
    
End Sub


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub


