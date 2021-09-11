VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderUpdate 
   Caption         =   "���� ����"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
   OleObjectBlob   =   "frmOrderUpdate.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmOrderUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bMatchedEstimateID As Boolean

Private Sub UserForm_Activate()
    Me.txtManagementID.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim order As Variant
    Dim db As Variant
    Dim contr As Control
    Dim orderId As Long
    Dim pos As Long
    Dim count As Long
    
    If clickOrderId <> "" Then              '�������� ���� ������Ȳ���� ����Ŭ���� ���
        If IsNumeric(clickOrderId) Then
            orderId = CLng(clickOrderId)
        Else
            orderId = clickOrderId
        End If
        clickOrderId = ""
    Else
        cRow = Selection.row                '���ְ���ȭ�鿡�� ����Ŭ������ ������ �� ��ȣ

        If cRow < 6 Or shtOrderAdmin.Range("B" & cRow).value = "" Then End         '�����Ͱ� �ִ� ���� �ƴ� ���� ����
        
        orderId = shtOrderAdmin.Cells(cRow, 2)
    End If
    
    'Label ��ġ ���߱�
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '�� ��ġ ����
    If orderUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderUpdateFormX
        Me.top = orderUpdateFormY
    End If
    
    '���� ������ �о����
    order = Get_Record_Array(shtOrder, orderId)
    
    Me.txtID.value = order(1)   'ID
    Me.txtManagementID.value = order(5) '������ȣ
    
    '������ȣ�� �������� ��������
    bMatchedEstimateID = False
    db = Get_DB(shtEstimate)
    db = Filtered_DB(db, Me.txtManagementID.value, 2, True)
    If Not IsEmpty(db) Then
        '������ ���� ��쿡�� �� ������ �������� ���
        count = UBound(db, 1)
        Me.txtEstimateID.value = db(count, 1)
        Me.txtEstimateCustomer.value = db(count, 4)
        Me.txtEstimateManager.value = db(count, 5)
        Me.txtEstimateName.value = db(count, 6)
    
        bMatchedEstimateID = True
    End If
    
    InitializeCboUnit
    InitializeOrderPayMethod
    InitializeOrderCategory
    InitializeLswCustomerAutoComplete
    
    Me.cboCategory.value = Trim(order(4))     '�з�
    Me.txtCustomer.value = order(6)     '�ŷ�ó
    Me.txtOrderName.value = order(7)    '���� ǰ��
    Me.txtMaterial.value = order(8)     '����
    Me.txtSize.value = order(9)             '�԰�
    Me.txtAmount.value = Format(order(10), "#,##0")   '����
    Me.cboUnit.value = Trim(order(11))            '����
    Me.txtUnitPrice.value = Format(order(12), "#,##0")     '�ܰ�
    Me.txtOrderPrice.value = Format(order(13), "#,##0")      '���ֱݾ�
    Me.txtWeight.value = order(14)          '�߷�
    Me.txtOrderDate.value = order(16)       '��������
    Me.txtDueDate.value = order(17)         '��������
    Me.txtReceivingDate.value = order(18)       '�԰�����
    Me.txtSpecificationDate.value = order(20)   '����
    Me.txtTaxInvoiceDate.value = order(21)      '��꼭
    Me.txtPaymentDate.value = order(22)     '��������
    Me.cboOrderPayMethod.value = Trim(order(24))       '��������
    Me.txtVAT.value = Format(order(25), "#,##0")             '�ΰ���
    
    Me.txtInsertDate.value = order(26)    '�������
    Me.txtUpdateDate.value = order(27)    '��������
    
    Me.txtMemo.value = order(29)            '�޸�
    Me.chkVAT.value = order(30)             '�ΰ��� ���� ����
    
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeOrderCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboCategory, db
End Sub

Sub InitializeOrderPayMethod()
    Dim db As Variant
    db = Get_DB(shtOrderPayMethod, True)

    Update_Cbo Me.cboOrderPayMethod, db
End Sub

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 108
        .Visible = False
    End With
End Sub

Sub UpdateOrder()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckOrderUpdateValidation = False Then
        Exit Sub
    End If

    '������ ������Ʈ
    Update_Record shtOrder, Me.txtID.value, _
        , , Me.cboCategory.value, _
        Me.txtManagementID.value, Me.txtCustomer.value, _
        Me.txtOrderName.value, Me.txtMaterial.value, _
        Me.txtSize.value, Me.txtAmount.value, _
        Me.cboUnit.value, Me.txtUnitPrice, _
        Me.txtOrderPrice.value, Me.txtWeight.value, _
        , Me.txtOrderDate.value, Me.txtDueDate.value, _
        Me.txtReceivingDate.value, , _
        Me.txtSpecificationDate.value, Me.txtTaxInvoiceDate.value, Me.txtPaymentDate.value, , _
        Me.cboOrderPayMethod.value, Me.txtVAT.value, _
        Me.txtInsertDate, Date, _
        Me.txtEstimateID.value, Me.txtMemo.value, Me.chkVAT.value

    Unload Me
    
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.InitializeLswOrderList
    Else
        shtOrderAdmin.Activate
        shtOrderAdmin.OrderSearch
        shtOrderAdmin.Range("K" & selectionRow).Select
    End If
    
End Sub


Function CheckOrderUpdateValidation()
    
    CheckOrderUpdateValidation = False
    
    'ǰ���� �ԷµǾ����� üũ
    If Trim(Me.txtOrderName.value) = "" Then
        MsgBox "ǰ���� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        Exit Function
    End If
    
    '������ȣ�� �ԷµǾ��� ��ȿ�� ������ȣ���� üũ
    If Trim(Me.txtManagementID.value) = "" Then
        MsgBox "������ȣ�� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        Exit Function
    End If
    
    If bMatchedEstimateID = False Then
        MsgBox "������ȣ�� ��ȿ���� �ʽ��ϴ�.", vbInformation, "�۾� Ȯ��"
        Exit Function
    End If
    
    CheckOrderUpdateValidation = True
    
End Function

Sub CalculateOrderUpdateCost()

    If Me.txtAmount.value = "" Then
        Me.txtOrderPrice.value = Me.txtUnitPrice.value
    Else
        If Me.txtUnitPrice.value = "" Then
            Me.txtOrderPrice.value = ""
        ElseIf IsNumeric(Me.txtUnitPrice.value) And IsNumeric(Me.txtAmount.value) Then
            Me.txtOrderPrice.value = Format(CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value), "#,##0")
        End If
    End If
    
    '�ΰ��� ���
    '���ݰ�꼭 ���ڰ� ���� ���, �ΰ��� ������ ��� �ΰ����� 0
    If Me.txtTaxInvoiceDate.value = "" Or chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        '�ΰ����� �ݾ��� 10%
        If Me.txtOrderPrice.value <> "" And Me.txtOrderPrice.value <> 0 Then
            Me.txtVAT.value = CLng(Me.txtOrderPrice.value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.value, "#,##0")
        End If
    End If

End Sub

Private Sub btnOrderUpdate_Click()
    UpdateOrder
End Sub

Private Sub btnOrderClose_Click()
    Unload Me
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = vbKeyReturn Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtOrderName.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If .ListItems.count = 1 Then
                If Me.txtCustomer.value <> .ListItems(1).Text Then
                    '��Ű�� ��� �ڵ��ϼ� ����� �Է°��� �ٸ��� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
                    .selectedItem = .ListItems(1)
                    .SetFocus
                Else
                    '�Է°��� �ڵ��ϼ� ����� ������ ���� �Է�ĭ���� �̵�
                    .Visible = False
                    Me.txtOrderName.SetFocus
                End If
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
            
            '���ְŷ�ó DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtOrderCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.value, 1, False)
            If IsEmpty(db) Then
                .Visible = False
            Else
                For i = 1 To UBound(db)
                    .ListItems.Add , , db(i, 1)
                    If i = 7 Then Exit For
                Next
            End If
            
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_DblClick()
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    With Me.lswCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtOrderName.SetFocus
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó ���� �� ����Ű ������ �� ���� �ŷ�ó�� �־��ְ� ��Ŀ���� ����(ǰ��)���� �̵�
    If KeyCode = vbKeyReturn Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtOrderName.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtOrderName_Enter()
    '�ڵ��ϼ� ����Ʈ���� ���ؼ� �Ѿ���� ���
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtReceivingDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateOrderUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub txtManagementID_AfterUpdate()
    Dim db As Variant
    
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
    
    Me.txtEstimateID.value = ""
    Me.txtEstimateCustomer.value = ""
    Me.txtEstimateManager.value = ""
    Me.txtEstimateName.value = ""
    
    '�Է��� ������ȣ�� �������̺��� �˻��ؼ� ����ID�� ������
    bMatchedEstimateID = False
    If Me.txtManagementID.value <> "" Then
        db = Get_DB(shtEstimate)
        db = Filtered_DB(db, Me.txtManagementID.value, 2, True)
        If IsEmpty(db) Then
            MsgBox "������ȣ�� �ش��ϴ� ����(����) ������ �����ϴ�.", vbInformation, "�۾� Ȯ��"
            Exit Sub
        Else
            If UBound(db, 1) = 1 Then
                Me.txtEstimateID.value = db(1, 1)
                Me.txtEstimateCustomer.value = db(1, 4)
                Me.txtEstimateManager.value = db(1, 5)
                Me.txtEstimateName.value = db(1, 6)
            
                bMatchedEstimateID = True
            Else
                MsgBox "������ȣ�� �������� ����(����) �������� ��� ���Դϴ�.", vbInformation, "�۾� Ȯ��"
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub txtAmount_AfterUpdate()
    Me.txtAmount.value = Trim(Me.txtAmount.value)
    
    If Me.txtAmount.value <> "" Then
         '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
            Exit Sub
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Me.txtUnitPrice.value = Trim(Me.txtUnitPrice.value)
    
    If Me.txtUnitPrice.value <> "" Then
        '�ܰ� ���� ���ڰ� �ƴ� ��� �����޽��� ���
        If IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = CLng(Me.txtUnitPrice.value)
        Else
            Me.txtUnitPrice.value = ""
            MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
            Exit Sub
        End If
    End If
    
    '�ݾ� 1,000�ڸ� �ĸ� ó��
    Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
End Sub

Private Sub txtOrderName_AfterUpdate()
    Me.txtOrderName.value = Trim(Me.txtOrderName.value)
End Sub

Private Sub txtMaterial_AfterUpdate()
    Me.txtMaterial.value = Trim(Me.txtMaterial.value)
End Sub

Private Sub txtOrderDate_AfterUpdate()
    Me.txtOrderDate.value = ConvertDateFormat(Me.txtOrderDate.value)
End Sub

Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub

Private Sub txtWeight_AfterUpdate()
    Me.txtWeight.value = Trim(Me.txtWeight.value)
End Sub

Private Sub txtReceivingDate_AfterUpdate()
    Me.txtReceivingDate.value = ConvertDateFormat(Me.txtReceivingDate.value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.value = ConvertDateFormat(Me.txtDueDate.value)
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = ConvertDateFormat(Me.txtPaymentDate.value)
End Sub

Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = ConvertDateFormat(Me.txtSpecificationDate.value)
End Sub

Private Sub txtTaxinvoiceDate_AfterUpdate()
    Me.txtTaxInvoiceDate.value = ConvertDateFormat(Me.txtTaxInvoiceDate.value)
   CalculateOrderUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateOrderUpdateCost
End Sub

Private Sub UserForm_Layout()
    orderUpdateFormX = Me.Left
    orderUpdateFormY = Me.top
End Sub

