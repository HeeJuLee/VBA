VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "���� ���"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "frmEstimateInsert.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub txtEstimateDate_Change()
    '���� �޽��� ����
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateID_AfterUpdate()
    '���� �޽��� ����
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    '���� �޽��� ����
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub UserForm_Initialize()

    '�ŷ�ó, ����� �޺��ڽ� ����
    InitializeCboCustomer
    InitializeCboUnit
    
    '������ �Է�â�� ��Ŀ��
    Me.txtEstimateName.SetFocus
    
End Sub

Private Sub cboCustomer_Change()
    '�޺��ڽ����� �ŷ�ó�� �����ϸ� �ش� �ŷ�ó�� ����ڷ� ����� �޺��ڽ��� ����
    InitializeCboManager
End Sub

Private Sub txtAmount_AfterUpdate()
    '�����޽��� ����
    Me.lblAmountError.Visible = False
    Me.lblInputFieldError.Visible = False
    
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
    Me.lblInputFieldError.Visible = False
    
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

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub


Private Sub btnEstimateInsert_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    InsertEstimate
End Sub

Private Sub btnEstimateClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub

'��ư ���콺���� ó��
'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnEstimateInsert_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEstimateInsert
End Sub

Private Sub btnEstimateInsert_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEstimateInsert
End Sub

Private Sub btnEstimateInsert_Enter()
OnHover_Css Me.btnEstimateInsert
End Sub

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnEstimateClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEstimateClose
End Sub

Private Sub btnEstimateClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEstimateClose
End Sub

Private Sub btnEstimateClose_Enter()
OnHover_Css Me.btnEstimateClose
End Sub

'�Ʒ� �ڵ带 �������� �߰��� ��, "btnXXX, btnYYY"�� ��ư�̸��� ��ǥ�� ������ ������ �����մϴ�.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnEstimateInsert, btnEstimateClose" ' ��ư �̸��� ��ǥ�� �����Ͽ� �Է��ϼ���.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
 For Each vList In vLists
 If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
 Next
Next
End Sub
'Ŀ�� �̵��� ��ư ������ �����ϴ� ������ɹ��� �������� �߰��մϴ�.
Private Sub OnHover_Css(lbl As Control): With lbl: .BackColor = RGB(211, 240, 224): .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl: .BackColor = &H8000000E: .BorderColor = -2147483638: End With: End Sub


Sub InitializeCboUnit()
    Dim DB As Variant
    DB = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, DB
End Sub

Sub InitializeCboCustomer()
    Dim DB As Variant
    DB = Get_DB(shtCustomer)

    Update_Cbo Me.cboCustomer, DB, 2
End Sub

Sub InitializeCboManager()
    Dim DB As Variant
    
    '����� DB�� �о�ͼ�
    DB = Get_DB(shtManager)
    '�ŷ�óID�� ���͸�
    DB = Filtered_DB(DB, Me.cboCustomer.Value, 2)
    
    '���� �޺��ڽ� ���������
    Me.cboManager.Clear
    
    '����ڰ� ������ �޺��ڽ��� �߰���
    If Not IsEmpty(DB) Then
        Update_Cbo Me.cboManager, DB, 3
    End If
End Sub



Sub InsertEstimate()
    Dim DB As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If CheckEstimateInsertValidation = False Then
        Exit Sub
    End If

    '�������� DB �о����
    DB = Get_DB(shtEstimate)
    
    '������ ������ȣ�� �ִ��� üũ
    blnUnique = IsUnique(DB, Me.txtEstimateID.Value, 3)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    Insert_Record shtEstimate, _
            Me.cboManager.Value, _
            Me.txtEstimateID.Value, _
            Me.txtLinkedID.Value, _
            Me.txtEstimateName.Value, _
            Me.txtSize.Value, _
            Me.txtAmount.Value, _
            Me.cboUnit.Value, _
            Me.txtUnitPrice.Value, _
            Me.txtEstimatePrice.Value, _
            Me.txtEstimateDate.Value, _
            , , , , _
            , , , , , , _
            , , , , _
            , , , , _
            Date, Date

    Unload Me
    
    shtEstimateAdmin.EstimateSearch
    
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
    If Trim(Me.txtEstimateID.Value) = "" Then
        bCorrect = False
        Me.lblEstimateIDEmpty.Visible = True
    Else
        Me.lblEstimateIDEmpty.Visible = False
    End If
    
    '������ �ԷµǾ����� üũ
    If Trim(Me.txtAmount.Value) = "" Then
        bCorrect = False
        Me.lblAmountEmpty.Visible = True
    Else
        Me.lblAmountEmpty.Visible = False
    End If
    
    '�����ܰ��� �ԷµǾ����� üũ
    If Trim(Me.txtUnitPrice.Value) = "" Then
        bCorrect = False
        Me.lblUnitPriceEmpty.Visible = True
    Else
        Me.lblUnitPriceEmpty.Visible = False
    End If
    
    '�������ڰ� �ԷµǾ����� üũ
    If Trim(Me.txtEstimateDate.Value) = "" Then
        bCorrect = False
        Me.lblEstimateDateEmpty.Visible = True
    Else
        Me.lblEstimateDateEmpty.Visible = False
    End If
    
    If bCorrect = False Then
        Me.lblInputFieldError.Visible = True
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
