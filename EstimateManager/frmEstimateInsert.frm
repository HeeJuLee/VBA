VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "���� ���"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
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
    Me.txtCustomer.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "Label*" Then
            contr.top = contr.top + 2
        End If
    Next

    '��Ʈ�� �ʱ�ȭ
    InitializeCboUnit
    InitializeLswCustomerAutoComplete
    InitializeLswManagerAutoComplete
    
    Me.txtEstimateDate.value = Date
    
    '�� ��ġ�� ����
    If estimateInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateInsertFormX
        Me.top = estimateInsertFormY
    End If
    
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
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
    blnUnique = IsUnique(db, Me.txtManagementID.value, 2)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation, "�۾� Ȯ��": Exit Sub
    
    Insert_Record shtEstimate, _
                  Me.txtManagementID.value, _
                  , _
                  Me.txtCustomer.value, _
                  Me.txtManager.value, _
                  Me.txtEstimateName.value, _
                  Me.txtSize.value, _
                  Me.txtAmount.value, _
                  Me.cboUnit.value, _
                  Me.txtUnitPrice.value, _
                  Me.txtEstimatePrice.value, _
                  Me.txtEstimateDate.value, _
                  , , , , _
                  , , , , , , _
                  Date, , _
                  , , _
                  , , , , _
                  , , False
            
    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    
End Sub

Function CheckEstimateInsertValidation()

    '�������� �ԷµǾ����� üũ
    If Trim(Me.txtEstimateName.value) = "" Then
        MsgBox "�������� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        CheckEstimateInsertValidation = False
        Me.txtEstimateName.SetFocus
        Exit Function
    End If
    
    '������ȣ�� �ԷµǾ����� üũ
    If Trim(Me.txtManagementID.value) = "" Then
        MsgBox "������ȣ�� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        CheckEstimateInsertValidation = False
        Me.txtManagementID.SetFocus
        Exit Function
    End If
    
    CheckEstimateInsertValidation = True
End Function

Sub CalculateEstimateInsertCost()

    '�������� �����̸� �ݾ��� �ܰ�
    If Me.txtAmount.value = "" Then
        Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Exit Sub
    End If
    
    '�ܰ��� ������ ���� ���� �ݾ����� ������
    If Me.txtUnitPrice.value <> "" And IsNumeric(Me.txtUnitPrice.value) Then
        Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")
    End If

End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnEstimateInsert_Click()
    InsertEstimate
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

Private Sub txtManager_Enter()
    '�ڵ��ϼ� ����Ʈ���� ���ؼ� �Ѿ���� ���
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .selectedItem.Text
            SetAutoManagementId .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtEstimateDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.btnEstimateInsert.SetFocus
    End If
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = vbKeyReturn Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtManager.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸�
            If .ListItems.count = 1 Then
                If Me.txtCustomer.value <> .ListItems(1).Text Then
                    '�ڵ��ϼ� ����� �Է°��� �ٸ��� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
                    .selectedItem = .ListItems(1)
                    .SetFocus
                Else
                    '�ڵ��ϼ� ����� �Է°��� ������ ���� �Է�ĭ���� �̵�
                    .Visible = False
                    Me.txtManager.SetFocus
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

Private Sub lswCustomerAutoComplete_DblClick()
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    With Me.lswCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
            SetAutoManagementId .selectedItem.Text
        End If
    End With
    
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó ���� �� ����Ű ������ �� ���� �ŷ�ó�� �־��ְ� ��Ŀ���� ����(�Ŵ�����)���� �̵�
    If KeyCode = vbKeyReturn Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                SetAutoManagementId .selectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
    
End Sub

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswManagerAutoComplete
        If KeyCode = vbKeyReturn Then
            '����Ű - ���� �Է�ĭ���� �̵�
            .Visible = False
            Me.txtEstimateName.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '��Ű�� ��쿡 �ڵ��ϼ� ����� �ϳ��̸�
            If .ListItems.count = 1 Then
                If Me.txtCustomer.value <> .ListItems(1).Text Then
                    '�ڵ��ϼ� ����� �Է°��� �ٸ��� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
                    .selectedItem = .ListItems(1)
                    .SetFocus
                Else
                    '�ڵ��ϼ� ����� �Է°��� ������ ���� �Է�ĭ���� �̵�
                    .Visible = False
                    Me.txtEstimateName.SetFocus
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

Private Sub lswManagerAutoComplete_DblClick()
    '����ڸ� ���� �־��ְ� ��Ŀ���� ������� �̵�
    With Me.lswManagerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtManager.value = .selectedItem.Text
            .Visible = False
            Me.txtEstimateName.SetFocus
        End If
    End With
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

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub txtAmount_AfterUpdate()
    
    If Me.txtAmount.value <> "" Then
        '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.value) Then
            MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
            Me.txtAmount.value = ""
            Exit Sub
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.value, "#,##0")
    
    CalculateEstimateInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.value <> "" Then
        '�ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
            Me.txtUnitPrice.value = ""
            Exit Sub
        End If
        
        '�ݾ� 1,000�ڸ� �ĸ� ó��
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    End If
    
    CalculateEstimateInsertCost
End Sub


Private Sub cboUnit_AfterUpdate()
    Me.cboUnit.value = Trim(Me.cboUnit.value)
End Sub


Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
    
    SetAutoManagementId Me.txtCustomer.value
End Sub

Sub SetAutoManagementId(customer)
    Dim db As Variant
    Dim day As String
    
    If customer <> "" Then
        db = Get_DB(shtEstimateCustomer, True)
        db = Filtered_DB(db, customer, 1, True)
        If isEmpty(db) Then
            Me.txtManagementID.value = Format(Date, "yy") & "Z" & Format(Date, "mmdd") & "-" & Format(time, "hhmm")
        Else
            Me.txtManagementID.value = Format(Date, "yy") & db(1, 2) & Format(Date, "mmdd") & "-" & Format(time, "hhmm")
        End If
    End If
End Sub

Private Sub txtEstimateDate_AfterUpdate()
    Me.txtEstimateDate.value = ConvertDateFormat(Me.txtEstimateDate.value)
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.txtEstimateName.value = Trim(Me.txtEstimateName.value)
End Sub


Private Sub txtManagementID_AfterUpdate()
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
End Sub

Private Sub txtManager_AfterUpdate()
    Me.txtManager.value = Trim(Me.txtManager.value)
End Sub

Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub

Private Sub UserForm_Layout()
    estimateInsertFormX = Me.Left
    estimateInsertFormY = Me.top
End Sub


