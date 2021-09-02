VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "���� ���"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   OleObjectBlob   =   "frmEstimateInsert.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    'Label ��ġ ���߱�
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '�� ��ġ ����
    If estimateInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateInsertFormX
        Me.top = estimateInsertFormY
    End If
    
    '�ŷ�ó, ����� �޺��ڽ� ����
    InitializeCboCustomer
    InitializeCboUnit
    
    '������ �Է�â�� ��Ŀ��
    Me.txtEstimateName.SetFocus
    
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtEstimateCustomer, True)

    Update_Cbo Me.cboCustomer, db, 1
End Sub

Sub InitializeCboManager()
    Dim db As Variant
    
    '����� DB�� �о�ͼ�
    db = Get_DB(shtEstimateManager, True)
    '�ŷ�ó������ ���͸�
    db = Filtered_DB(db, Me.cboCustomer.Value, 1, True)
    
    '���� �޺��ڽ� ���������
    Me.cboManager.Clear
    
    '����ڰ� ������ �޺��ڽ��� �߰���
    If Not IsEmpty(db) Then
        Update_Cbo Me.cboManager, db, 2
    End If
End Sub

Sub InsertEstimate()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckEstimateInsertValidation = False Then
        Exit Sub
    End If

    '�������� DB �о����
    db = Get_DB(shtEstimate)
    
    '������ ������ȣ�� �ִ��� üũ
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 3)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    Insert_Record shtEstimate, _
            Me.txtManagementID.Value, _
            Me.txtLinkedID.Value, _
            Me.cboCustomer.Value, _
            Me.cboManager.Value, _
            Me.txtEstimateName.Value, _
            Me.txtSize.Value, _
            Me.txtAmount.Value, _
            Me.cboUnit.Value, _
            Me.txtUnitPrice.Value, _
            Me.txtEstimatePrice.Value, _
            Me.txtEstimateDate.Value, _
            , , , , _
            , , , , , , _
            Date, , _
            , , , , , , , , False
            
    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.GoToEnd
    
End Sub

Function CheckEstimateInsertValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '�������� �ԷµǾ����� üũ
    If Trim(Me.txtEstimateName.Value) = "" Then
        bCorrect = False
        Me.lblEstimateNameEmpty.Visible = True
    Else
        Me.lblEstimateNameEmpty.Visible = False
    End If
    
    '������ȣ�� �ԷµǾ����� üũ
    If Trim(Me.txtManagementID.Value) = "" Then
        bCorrect = False
        Me.lblManagementIDEmpty.Visible = True
    Else
        Me.lblManagementIDEmpty.Visible = False
    End If
    
    CheckEstimateInsertValidation = bCorrect
End Function

Sub CalculateEstimateInsertCost()

    '�������� �����̸� �����ݾ��� �����ܰ�
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
        Exit Sub
    End If
    
    '�����ܰ��� ������ ���� ���� �����ݾ����� ������
    If Me.txtUnitPrice.Value <> "" And IsNumeric(Me.txtUnitPrice.Value) Then
        Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")
    End If

End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnEstimateInsert_Click()
    InsertEstimate
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub cboCustomer_Change()
    '�޺��ڽ����� �ŷ�ó�� �����ϸ� �ش� �ŷ�ó�� ����ڷ� ����� �޺��ڽ��� ����
    InitializeCboManager
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.lblEstimateNameEmpty.Visible = False
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.lblManagementIDEmpty.Visible = False
End Sub

Private Sub txtAmount_AfterUpdate()
    '�����޽��� ����
    Me.lblAmountError.Visible = False
'    Me.lblInputFieldError.Visible = False
    
    If Me.txtAmount.Value <> "" Then
         '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            Me.lblAmountError.Visible = True
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    CalculateEstimateInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
     '�����޽��� ����
    Me.lblUnitPriceError.Visible = False
    'Me.lblInputFieldError.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '�����ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '�ݾ� 1,000�ڸ� �ĸ� ó��
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    End If
    
    CalculateEstimateInsertCost
End Sub

Private Sub UserForm_Layout()
    estimateInsertFormX = Me.Left
    estimateInsertFormY = Me.top
End Sub
