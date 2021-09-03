VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "���� ���"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   OleObjectBlob   =   "frmEstimateInsert.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Me.txtEstimateName.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    'Label ��ġ ���߱�
'    For Each contr In Me.Controls
'    If contr.Name Like "Label*" Then
'        contr.top = contr.top + 2
'    End If
'    Next

    '��Ʈ�� �ʱ�ȭ
    InitializeCboUnit
    InitializeLswCustomerAutoComplete
    InitializeLswManagerAutoComplete
    
    Me.txtEstimateDate.Value = Date
    
    '�� ��ġ�� ����
    If estimateInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateInsertFormX
        Me.top = estimateInsertFormY
    End If
    
    '�ؽ�Ʈ�ڽ� ���̺� ���� �� ���ڻ� ����
    Me.lblEstimateName.BackColor = RGB(84, 130, 53)
    Me.lblManagementID.BackColor = RGB(84, 130, 53)
    Me.lblLinkedID.BackColor = RGB(48, 84, 150)
    Me.lblCustomer.BackColor = RGB(48, 84, 150)
    Me.lblManager.BackColor = RGB(48, 84, 150)
    Me.lblSize.BackColor = RGB(48, 84, 150)
    Me.lblAmount.BackColor = RGB(48, 84, 150)
    Me.lblUnit.BackColor = RGB(48, 84, 150)
    Me.lblUnitPrice.BackColor = RGB(48, 84, 150)
    Me.lblEstimatePrice.BackColor = RGB(48, 84, 150)
    Me.lblEstimateDate.BackColor = RGB(48, 84, 150)
    
    Me.lblEstimateName.ForeColor = RGB(255, 255, 255)
    Me.lblManagementID.ForeColor = RGB(255, 255, 255)
    Me.lblLinkedID.ForeColor = RGB(255, 255, 255)
    Me.lblCustomer.ForeColor = RGB(255, 255, 255)
    Me.lblManager.ForeColor = RGB(255, 255, 255)
    Me.lblSize.ForeColor = RGB(255, 255, 255)
    Me.lblAmount.ForeColor = RGB(255, 255, 255)
    Me.lblUnit.ForeColor = RGB(255, 255, 255)
    Me.lblUnitPrice.ForeColor = RGB(255, 255, 255)
    Me.lblEstimatePrice.ForeColor = RGB(255, 255, 255)
    Me.lblEstimateDate.ForeColor = RGB(255, 255, 255)
    
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
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
            Trim(Me.txtManagementID.Value), _
            Trim(Me.txtLinkedID.Value), _
            Trim(Me.txtCustomer.Value), _
            Trim(Me.txtManager.Value), _
            Trim(Me.txtEstimateName.Value), _
            Trim(Me.txtSize.Value), _
            Trim(Me.txtAmount.Value), _
            Trim(Me.cboUnit.Value), _
            Trim(Me.txtUnitPrice.Value), _
            Trim(Me.txtEstimatePrice.Value), _
            Trim(Me.txtEstimateDate.Value), _
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

    '�������� �ԷµǾ����� üũ
    If Trim(Me.txtEstimateName.Value) = "" Then
        MsgBox "�������� �Է��ϼ���."
        CheckEstimateInsertValidation = False
        Me.txtEstimateName.SetFocus
        Exit Function
    End If
    
    '������ȣ�� �ԷµǾ����� üũ
    If Trim(Me.txtManagementID.Value) = "" Then
        MsgBox "������ȣ�� �Է��ϼ���."
        CheckEstimateInsertValidation = False
        Me.txtManagementID.SetFocus
        Exit Function
    End If
    
    CheckEstimateInsertValidation = True
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

Private Sub txtEstimateDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnEstimateInsert.SetFocus
    End If
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '����Ű - ���� �Է�ĭ���� �̵�
        Me.lswCustomerAutoComplete.Visible = False
        Me.txtManager.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '��Ű, �Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
        With Me.lswCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '�ŷ�ó �ڵ��ϼ� ó��
    With Me.lswCustomerAutoComplete
        If Me.txtCustomer.Value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '�����ŷ�ó DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtEstimateCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.Value, 1, False)
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

Private Sub lswCustomerAutoComplete_DblClick()
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    With Me.lswCustomerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtCustomer.Value = .SelectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó ���� �� ����Ű ������ �� ���� �ŷ�ó�� �־��ְ� ��Ŀ���� ����(�Ŵ�����)���� �̵�
    If KeyCode = 13 Then
        With Me.lswCustomerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtCustomer.Value = .SelectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '����Ű - ���� �Է�ĭ���� �̵�
        Me.lswManagerAutoComplete.Visible = False
        Me.txtSize.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '��Ű, �Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
        With Me.lswManagerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManager_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '����� �ڵ��ϼ� ó��
    With Me.lswManagerAutoComplete
        If Me.txtManager.Value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '��������� DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtEstimateManager, True)
            db = Filtered_DB(db, Me.txtManager.Value, 1, False)
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

Private Sub lswManagerAutoComplete_DblClick()
    '����ڸ� ���� �־��ְ� ��Ŀ���� ������� �̵�
    With Me.lswManagerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtManager.Value = .SelectedItem.Text
            .Visible = False
            Me.txtSize.SetFocus
        End If
    End With
End Sub

Private Sub lswManagerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '����� ���� �� ����Ű ������ �� ���� ����ڸ� �־��ְ� ��Ŀ���� ����(������)���� �̵�
    If KeyCode = 13 Then
        With Me.lswManagerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtManager.Value = .SelectedItem.Text
                .Visible = False
                Me.txtSize.SetFocus
            End If
        End With
    End If
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub txtAmount_AfterUpdate()
    
    If Me.txtAmount.Value <> "" Then
         '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.Value) Then
            MsgBox "���ڸ� �Է��ϼ���."
            Me.txtAmount.Value = ""
            Exit Sub
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    CalculateEstimateInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.Value <> "" Then
        '�����ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            MsgBox "���ڸ� �Է��ϼ���."
            Me.txtUnitPrice.Value = ""
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
