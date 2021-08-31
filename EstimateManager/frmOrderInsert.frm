VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderInsert 
   Caption         =   "���� ���"
   ClientHeight    =   8775.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "frmOrderInsert.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmOrderInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bMatchedEstimateID As Boolean

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    'Label ��ġ ���߱�
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '�� ��ġ ����
    If orderInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderInsertFormX
        Me.top = orderInsertFormY
    End If
    
    InitializeOrderCategory
    InitializeCboUnit
    
    '���ָ� �Է�â�� ��Ŀ��
    Me.txtOrderName.SetFocus
    
    bMatchedEstimateID = False
    
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

Sub InsertOrder()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckOrderInsertValidation = False Then
        Exit Sub
    End If

    Insert_Record shtOrder, _
            , Me.cboCategory.Value, Me.txtManagementID.Value, _
            Me.txtCustomer.Value, _
            Me.txtOrderName.Value, _
            Me.txtMaterial.Value, _
            Me.txtSize.Value, _
            Me.txtAmount.Value, _
            Me.cboUnit.Value, _
            Me.txtUnitPrice.Value, _
            Me.txtOrderPrice.Value, _
            Me.txtWeight.Value, _
            Me.txtOrderDate.Value, _
            , , , _
            , , , , , _
            Date, , Me.txtEstimateID, False
            
    Unload Me
    
    shtOrderAdmin.OrderSearch
    shtOrderAdmin.GoToEnd
    
End Sub


Function CheckOrderInsertValidation()
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
    
    CheckOrderInsertValidation = bCorrect
End Function

Sub CalculateOrderInsertCost()

    '�������� �����̸� ���ֱݾ��� �ܰ�
    If Me.txtAmount.Value = "" Then
        Me.txtOrderPrice.Value = Me.txtUnitPrice.Value
        Exit Sub
    End If
    
    '�ܰ��� ������ ���� ���� ���ֱݾ����� ������
    If Me.txtUnitPrice.Value <> "" And IsNumeric(Me.txtUnitPrice.Value) Then
        Me.txtOrderPrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        Me.txtOrderPrice.Text = Format(Me.txtOrderPrice.Value, "#,##0")
    End If

End Sub

Private Sub btnOrderClose_Click()
    Unload Me
    
    shtOrderAdmin.OrderSearch
End Sub

Private Sub btnOrderInsert_Click()
    InsertOrder
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub txtOrderName_AfterUpdate()
    Me.lblOrderNameEmpty.Visible = False
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
    '�����޽��� ����
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
    
    CalculateOrderInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
     '�����޽��� ����
    Me.lblUnitPriceError.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '�����ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = CLng(Me.txtUnitPrice.Value)
        Else
            Me.txtUnitPrice.Value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '�ݾ� 1,000�ڸ� �ĸ� ó��
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    End If
    
    CalculateOrderInsertCost
End Sub


Private Sub UserForm_Layout()
    orderInsertFormX = Me.Left
    orderInsertFormY = Me.top
End Sub


