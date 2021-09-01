VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderUpdate 
   Caption         =   "���� ����"
   ClientHeight    =   9435.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
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

Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim order As Variant
    Dim db As Variant
    Dim contr As Control
    Dim orderId As Long
    Dim pos As Long
    
    If clickOrderId <> "" Then              '�������� ���� ������Ȳ���� ����Ŭ���� ���
        If IsNumeric(clickOrderId) Then
            orderId = CLng(clickOrderId)
        Else
            orderId = clickOrderId
        End If
        clickOrderId = ""
    Else
        cRow = Selection.row                '���ְ���ȭ�鿡�� ����Ŭ������ ������ �� ��ȣ

        If cRow < 6 Or shtOrderAdmin.Range("B" & cRow).Value = "" Then End         '�����Ͱ� �ִ� ���� �ƴ� ���� ����
        
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
    
    Me.txtID.Value = order(1)   'ID
    Me.txtOrderName.Value = order(6)    '���� ǰ��
    Me.txtManagementID.Value = order(4) '������ȣ
    
    '������ȣ�� �������� ��������
    bMatchedEstimateID = False
    db = Get_DB(shtEstimate)
    db = Filtered_DB(db, Me.txtManagementID.Value, 2, True)
    If IsEmpty(db) Then
        Me.lblManagementIDError.Caption = "������ȣ ����"
        Me.lblManagementIDError.Visible = True
    Else
        If UBound(db, 1) = 1 Then
            Me.txtEstimateID.Value = db(1, 1)
            Me.txtEstimateCustomer.Value = db(1, 4)
            Me.txtEstimateManager.Value = db(1, 5)
            Me.txtEstimateName.Value = db(1, 6)
        
            bMatchedEstimateID = True
        Else
            Me.lblManagementIDError.Caption = "������ȣ �ߺ�"
            Me.lblManagementIDError.Visible = True
        End If
    End If
    
    InitializeOrderCategory
    InitializeCboUnit
    InitializePayMethod
    
    Me.cboCategory.Value = Trim(order(3))     '�з�
    Me.txtCustomer.Value = order(5)     '�ŷ�ó
    Me.txtMaterial.Value = order(7)     '����
    Me.txtSize.Value = order(8)             '�԰�
    Me.txtAmount.Value = Format(order(9), "#,##0")   '����
    Me.cboUnit.Value = Trim(order(10))            '����
    Me.txtUnitPrice.Value = Format(order(11), "#,##0")     '�ܰ�
    Me.txtOrderPrice.Value = Format(order(12), "#,##0")      '���ֱݾ�
    Me.txtWeight.Value = order(13)          '�߷�
    Me.txtOrderDate.Value = order(14)       '��������
    Me.txtDueDate.Value = order(15)         '��������
    Me.txtDeliveryDate.Value = order(16)       '�԰�����
    Me.txtSpecificationDate.Value = order(17)   '����
    Me.txtTaxInvoiceDate.Value = order(18)      '��꼭
    Me.txtPaymentDate.Value = order(19)     '��������
    Me.cboPayMethod.Value = Trim(order(21))       '��������
    Me.txtVAT.Value = Format(order(22), "#,##0")             '�ΰ���
    
    Me.txtInsertDate.Value = order(23)    '�������
    Me.txtUpdateDate.Value = order(24)    '��������
    
    Me.txtMemo.Value = order(26)            '�޸�
    Me.chkVAT.Value = order(27)             '�ΰ��� ���� ����
    
    '���ָ� �Է�â�� ��Ŀ��
    Me.txtOrderName.SetFocus
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

Sub InitializePayMethod()
    Dim db As Variant
    db = Get_DB(shtPayMethod, True)

    Update_Cbo Me.cboPayMethod, db
End Sub


Sub UpdateOrder()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckOrderUpdateValidation = False Then
        Exit Sub
    End If

    '������ ������Ʈ
    Update_Record shtOrder, Me.txtID.Value, _
        , Me.cboCategory.Value, _
        Me.txtManagementID.Value, Me.txtCustomer.Value, _
        Me.txtOrderName.Value, Me.txtMaterial.Value, _
        Me.txtSize.Value, Me.txtAmount.Value, _
        Me.cboUnit.Value, Me.txtUnitPrice, _
        Me.txtOrderPrice.Value, Me.txtWeight.Value, _
        Me.txtOrderDate.Value, Me.txtDueDate.Value, _
        Me.txtDeliveryDate.Value, Me.txtSpecificationDate.Value, _
        Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, _
        , Me.cboPayMethod.Value, Me.txtVAT.Value, _
        Me.txtInsertDate, Date, _
        Me.txtEstimateID.Value, Me.txtMemo.Value, Me.chkVAT.Value

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
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '���ָ��� �ԷµǾ����� üũ
    If Trim(Me.txtOrderName.Value) = "" Then
        bCorrect = False
        Me.lblOrderNameEmpty.Visible = True
    Else
        Me.lblOrderNameEmpty.Visible = False
    End If
    
    '������ȣ�� �ԷµǾ��� ��ȿ�� ������ȣ���� üũ
    If Trim(Me.txtManagementID.Value) = "" Or bMatchedEstimateID = False Then
        bCorrect = False
        Me.lblManagementIDEmpty.Visible = True
    Else
        Me.lblManagementIDEmpty.Visible = False
    End If
    
    CheckOrderUpdateValidation = bCorrect
End Function

Sub CalculateOrderUpdateCost()

    '���ֱݾ� ���
    If Me.txtUnitPrice.Value <> "" Then
        If Me.txtAmount.Value = "" Then
            Me.txtOrderPrice.Value = Me.txtUnitPrice.Value
        Else
            Me.txtOrderPrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        End If
    End If
    Me.txtOrderPrice.Text = Format(Me.txtOrderPrice.Value, "#,##0")
    
    '�ΰ��� ���
    '���ݰ�꼭 ���ڰ� ���� ���, �ΰ��� ������ ��� �ΰ����� 0
    If Me.txtTaxInvoiceDate.Value = "" Or chkVAT.Value = True Then
        Me.txtVAT.Value = 0
    Else
        '�ΰ����� �ݾ��� 10%
        If Me.txtOrderPrice.Value <> "" And Me.txtOrderPrice.Value <> 0 Then
            Me.txtVAT.Value = CLng(Me.txtOrderPrice.Value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.Value, "#,##0")
        End If
    End If

End Sub

Private Sub btnOrderUpdate_Click()
    UpdateOrder
End Sub

Private Sub btnOrderClose_Click()
    Unload Me
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateOrderUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub txtManagementID_AfterUpdate()
    Dim db As Variant
    
    Me.lblManagementIDEmpty.Visible = False
    Me.lblManagementIDError.Visible = False
    
    Me.txtEstimateID.Value = ""
    Me.txtEstimateCustomer.Value = ""
    Me.txtEstimateManager.Value = ""
    Me.txtEstimateName.Value = ""
    
    '�Է��� ������ȣ�� �������̺��� �˻��ؼ� ����ID�� ������
    bMatchedEstimateID = False
    If Me.txtManagementID.Value <> "" Then
        db = Get_DB(shtEstimate)
        db = Filtered_DB(db, Me.txtManagementID.Value, 2, True)
        If IsEmpty(db) Then
            Me.lblManagementIDError.Caption = "������ȣ ����"
            Me.lblManagementIDError.Visible = True
        Else
            If UBound(db, 1) = 1 Then
                Me.txtEstimateID.Value = db(1, 1)
                Me.txtEstimateCustomer.Value = db(1, 4)
                Me.txtEstimateManager.Value = db(1, 5)
                Me.txtEstimateName.Value = db(1, 6)
            
                bMatchedEstimateID = True
            Else
                Me.lblManagementIDError.Caption = "������ȣ �ߺ�"
                Me.lblManagementIDError.Visible = True
            End If
        End If
    End If
    
End Sub

Private Sub txtAmount_AfterUpdate()
    Me.lblAmountError.Visible = False
    
    If Me.txtAmount.Value <> "" Then
         '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            Me.lblAmountError.Visible = True
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Me.lblUnitPriceError.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '�����ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '�ܰ� 1,000�ڸ� �ĸ� ó��
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    End If
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateOrderUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateOrderUpdateCost
End Sub

Private Sub UserForm_Layout()
    orderUpdateFormX = Me.Left
    orderUpdateFormY = Me.top
End Sub

