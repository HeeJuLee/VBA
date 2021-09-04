VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "���� ����"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19200
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




Private Sub cboUnit_AfterUpdate()
    Me.cboUnit.Value = Trim(Me.cboUnit.Value)
End Sub


Private Sub txtBidDate_AfterUpdate()
    Me.txtBidDate.Value = Trim(Me.txtBidDate.Value)
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.Value = Trim(Me.txtCustomer.Value)
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    Me.txtDeliveryDate.Value = Trim(Me.txtDeliveryDate.Value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.Value = Trim(Me.txtDueDate.Value)
End Sub


Private Sub txtEstimateDate_AfterUpdate()
    Me.txtEstimateDate.Value = Trim(Me.txtEstimateDate.Value)
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    Me.txtInsuranceDate.Value = Trim(Me.txtInsuranceDate.Value)
End Sub

Private Sub txtManager_AfterUpdate()
    Me.txtManager.Value = Trim(Me.txtManager.Value)
End Sub


Private Sub txtSize_AfterUpdate()
    Me.txtSize.Value = Trim(Me.txtSize.Value)
End Sub


Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.Value = Trim(Me.txtSpecificationDate.Value)
End Sub

Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim db As Variant
    Dim contr As Control
    
    If clickEstimateId <> "" Then              '���ְ������� ����Ŭ���� ���
        currentEstimateId = CLng(clickEstimateId)
        clickEstimateId = ""
    Else
        '������ �� ��ȣ
        cRow = Selection.row
    
        '�����Ͱ� �ִ� ���� �ƴ� ���� ����
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then
            MsgBox "������ ���� ���� ���� ������ �� �������� ��ư�� Ŭ���ϼ���."
            End
        End If
        
        currentEstimateId = shtEstimateAdmin.Cells(cRow, 2)
    End If
    
     '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            'contr.top = contr.top + 2
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

    Me.txtID.Value = estimate(1)    'ID
    Me.txtEstimateName.Value = estimate(6)  '������
    Me.txtManagementID.Value = estimate(2)    '������ȣ
    Me.txtLinkedID.Value = estimate(3)  '�����ȣ
    
    Me.txtCustomer = estimate(4)   '�ŷ�ó
    Me.txtManager = estimate(5)   '�����
    
    Me.txtSize.Value = estimate(7)  '�԰�
    Me.txtAmount.Value = Format(estimate(8), "#,##0")   '����
    InitializeCboUnit
    Me.cboUnit.Value = Trim(estimate(9))  '����, ID�� �����Ƿ� ���� value ������ ���õ�
    Me.txtUnitPrice.Value = Format(estimate(10), "#,##0")     '�����ܰ�
    Me.txtEstimatePrice.Value = Format(estimate(11), "#,##0")     '�����ݾ�
    
    Me.txtEstimateDate.Value = estimate(12)    '��������
    Me.txtBidDate.Value = estimate(13)    '��������
    Me.txtAcceptedDate.Value = estimate(14)    '��������
    Me.txtDeliveryDate.Value = estimate(15)    '��ǰ����
    Me.txtInsuranceDate.Value = estimate(16)    '��������
    
    Me.txtProductionTotalCost.Value = Format(estimate(17), "#,##0")   '���డ
    Me.txtBidPrice.Value = Format(estimate(18), "#,##0")    '������
    Me.txtBidMargin.Value = Format(estimate(19), "#,##0")    '����
    Me.txtBidMarginRate.Value = Format(estimate(20), "0.0%")    '������
    Me.txtAcceptedPrice.Value = Format(estimate(21), "#,##0")    '���ֱݾ�
    Me.txtAcceptedMargin.Value = Format(estimate(22), "#,##0")   '��������
    
    Me.txtInsertDate.Value = estimate(23)    '�������
    Me.txtUpdateDate.Value = estimate(24)    '��������
    
    InitializeCboCategory
    Me.cboCategory.Value = Trim(estimate(25))   '�з�
    '26�� ������
    Me.txtSpecificationDate.Value = estimate(27)    '�ŷ�����
    Me.txtTaxInvoiceDate.Value = estimate(28)    '���ݰ�꼭
    Me.txtPaymentDate.Value = estimate(29)    '��������
    Me.txtExpectPaymentDate.Value = estimate(30)  '���������
    Me.txtExpectPaymentMonth.Value = Format(estimate(30), "mm" & "��")  '���������
    Me.txtVAT.Value = Format(estimate(31), "#,##0")    '�ΰ���
    Me.txtMemo.Value = estimate(32)
    Me.chkVAT.Value = estimate(33)      '�ΰ��� ���� ����
    
    Me.txtPaid.Value = Format(estimate(34), "#,##0")      '�Աݾ�
    Me.txtRemaining.Value = Format(estimate(35), "#,##0")      '���Աݾ�
    Me.chkDividePay.Value = estimate(36)      '���Ұ��� ����
    
    '���� �� ������ȣ
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '���� ��Ȳ
    
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
    db = Filtered_DB(db, Me.cboCustomer.Value, 1, True)
    
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


Sub InitializeLswOrderCustomerAutoComplete()
    
    With Me.lswOrderCustomerAutoComplete
        .View = lvwList
'        .Gridlines = True
'        .FullRowSelect = True
'        .HideColumnHeaders = False
'        .HideSelection = True
'        .FullRowSelect = True
'        .MultiSelect = False
'        .LabelEdit = lvwManual
        .Height = 126
        .Visible = False
    End With
End Sub


Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� ���� ������ �о��
    db = Get_DB(shtOrder)
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.Value, 28, True)
    End If
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, "<>" & "����", 4)
    End If
    
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
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 70
        .ColumnHeaders.Add , , "����", 62
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
        
        .ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 7))   'ǰ��
                li.ListSubItems.Add , , db(i, 1)        'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_����
                li.ListSubItems.Add , , db(i, 5)        '������ȣ
                li.ListSubItems.Add , , db(i, 4)        '�з�
                li.ListSubItems.Add , , db(i, 6)        '�ŷ�ó
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
                li.Selected = False
                
                If IsNumeric(db(i, 13)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 13))
                End If
            Next
        End If
        
        If totalCost <> 0 Then
            Me.txtExecutionCost.Value = Format(totalCost, "#,##0")
        End If
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
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    '������ ������Ʈ
    Update_Record shtEstimate, Me.txtID.Value, _
        Me.txtManagementID.Value, Me.txtLinkedID.Value, _
        Me.txtCustomer.Value, Me.txtManager.Value, _
        Me.txtEstimateName.Value, Me.txtSize.Value, _
        Me.txtAmount.Value, Me.cboUnit.Value, _
        Me.txtUnitPrice.Value, Me.txtEstimatePrice.Value, _
        Me.txtEstimateDate.Value, Me.txtBidDate.Value, _
        Me.txtAcceptedDate.Value, Me.txtDeliveryDate.Value, _
        Me.txtInsuranceDate.Value, Me.txtProductionTotalCost.Value, _
        Me.txtBidPrice.Value, Me.txtBidMargin.Value, _
        Me.txtBidMarginRate.Value, Me.txtAcceptedPrice.Value, _
        Me.txtAcceptedMargin.Value, _
        Me.txtInsertDate.Value, Date, _
        Me.cboCategory.Value, Me.txtDueDate.Value, _
        Me.txtSpecificationDate.Value, Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, Me.txtExpectPaymentDate.Value, _
        Me.txtVAT.Value, Me.txtMemo.Value, Me.chkVAT.Value, _
        Me.txtPaid.Value, Me.txtRemaining.Value, Me.chkDividePay
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.Range("H" & selectionRow).Select
End Sub

Function CheckEstimateUpdateValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '�������� �ԷµǾ����� üũ
    If Me.txtEstimateName.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "�������� �Է��ϼ���."
    End If
    
    '������ȣ�� �ԷµǾ����� üũ
    If Me.txtManagementID.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "������ȣ�� �Է��ϼ���."
    End If
    
    If bCorrect = False Then
        Me.lblErrorMessage.Visible = True
    Else
        Me.lblErrorMessage.Visible = False
    End If
    
    CheckEstimateUpdateValidation = bCorrect
End Function

Sub CalculateEstimateUpdateCost()

    '�����ݾ� ���
    '�������� �����̸� �����ݾ��� �����ܰ�
    If Me.txtUnitPrice <> "" Then
        If Me.txtAmount.Value = "" Then
            Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
        Else
            Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")

    '�������װ� �������� ���
    If Me.txtBidPrice.Value <> "" And Me.txtProductionTotalCost.Value <> "" Then
        '�������� = ������ - ������డ
        Me.txtBidMargin.Value = Format(CLng(Me.txtBidPrice.Value) - CLng(Me.txtProductionTotalCost.Value), "#,##0")
        '�������� = �������� / ������
        If Me.txtBidPrice.Value <> "0" Then
            Me.txtBidMarginRate.Value = Format(CLng(Me.txtBidMargin.Value) / CLng(Me.txtBidPrice.Value), "0.0%")
        End If
    Else
        Me.txtBidMargin.Value = 0
    End If

    '��������, ������ ���
    If Me.txtAcceptedPrice.Value <> "" And Me.txtExecutionCost.Value <> "" Then
        '�������� = ���ֱݾ� - ���డ
        Me.txtAcceptedMargin.Value = Format(CLng(Me.txtAcceptedPrice.Value) - CLng(Me.txtExecutionCost.Value), "#,##0")
        '������ = �������� / ���ֱݾ�
        If Me.txtAcceptedPrice.Value <> "0" Then
            Me.txtAcceptedMarginRate.Value = Format(CLng(Me.txtAcceptedMargin.Value) / CLng(Me.txtAcceptedPrice.Value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.Value = ""
        Me.txtAcceptedMarginRate.Value = ""
    End If

    '�ΰ��� ���
    '���ݰ�꼭 ���ڰ� ���� ���, �ΰ��� ������ ��� �ΰ����� 0
    If Me.txtTaxInvoiceDate.Value = "" Or chkVAT.Value = True Then
        Me.txtVAT.Value = 0
    Else
        '�ΰ����� ���ֱݾ��� 10%
        If Me.txtAcceptedPrice.Value <> "" And Me.txtAcceptedPrice.Value <> 0 Then
            Me.txtVAT.Value = CLng(Me.txtAcceptedPrice.Value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.Value, "#,##0")
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


Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2, True)
    
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


Private Sub lswOrderList_DblClick()
    With Me.lswOrderList
        If Not .SelectedItem Is Nothing Then
            clickOrderId = .SelectedItem.ListSubItems(1)
            
            If frmOrderUpdate.Visible = True Then
                Unload frmOrderUpdate
            End If
        
            frmOrderUpdate.Show (False)
        End If
    End With
End Sub

Private Sub btnEstimateUpdate_Click()
    UpdateEstimate
End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnOrderListDelete_Click()
    Dim li As ListItem
    Dim count As Long
    Dim YN As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "������ ���ָ� �����ϼ���.": Exit Sub
    
    YN = MsgBox("������ " & count & "�� ���ָ� �����մϴ�.", vbYesNo)
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            '���� ���̺��� ����
            Delete_Record shtOrder, li.SubItems(1)
        End If
    Next
    
    If count > 0 Then
        InitializeLswOrderList
    End If
End Sub

Private Sub btnProduction_Click()
    If isFormLoaded("frmProduction") Then
        Unload frmProduction
    End If
    frmProduction.Show (False)
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub


Private Sub btnPayHistoryInsert_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnEstimateUpdate.SetFocus
    End If
End Sub

Private Sub btnProduction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub

Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub

Private Sub lswOrderCustomerAutoComplete_DblClick()
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    With Me.lswOrderCustomerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtProductionCustomer.Value = .SelectedItem.Text
            .Visible = False
            Me.txtProductionItem.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    If KeyCode = 13 Then
        With Me.lswOrderCustomerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtProductionCustomer.Value = .SelectedItem.Text
                .Visible = False
                Me.txtProductionItem.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "��")
End Sub

Private Sub cboCustomer_Change()
    InitializeCboManager
End Sub

Private Sub txtEstimateDate_Change()
    
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.txtManagementID.Value = Trim(Me.txtManagementID.Value)
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.txtEstimateName.Value = Trim(Me.txtEstimateName.Value)
End Sub

Private Sub txtAmount_AfterUpdate()

    If Me.txtAmount.Value <> "" Then
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '���� 1,000�ڸ� �ĸ� ó��
            Me.txtAmount.Value = Format(Me.txtAmount.Value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.Value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�����ܰ� 1,000�ڸ� �ĸ� ó��
            Me.txtUnitPrice.Value = Format(Me.txtUnitPrice.Value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtBidPrice_AfterUpdate()
    
    If Me.txtBidPrice.Value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtBidPrice.Value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '�����ݾ� 1,000�ڸ� �ĸ� ó��
            Me.txtBidPrice.Value = Format(Me.txtBidPrice.Value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If Me.txtAcceptedPrice.Value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.Value) Then
            Me.txtAcceptedPrice.Value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            Me.txtAcceptedPrice.Value = Format(Me.txtAcceptedPrice.Value, "#,##0")
            
            CalculateEstimateUpdateCost
        End If
    End If
End Sub

Private Sub txtProductionTotalCost_AfterUpdate()
    
    If Me.txtProductionTotalCost.Value <> "" Then
        If Not IsNumeric(Me.txtProductionTotalCost.Value) Then
            Me.txtProductionTotalCost.Value = ""
            MsgBox "���ڸ� �Է��ϼ���."
        Else
            '������డ 1,000�ڸ� �ĸ� ó��
            Me.txtProductionTotalCost.Value = Format(Me.txtProductionTotalCost.Value, "#,##0")
            
            '��� �ʵ� ���
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtExecutionCost_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    Me.txtAcceptedDate.Value = Trim(Me.txtAcceptedDate.Value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
    Me.txtTaxInvoiceDate.Value = Trim(Me.txtTaxInvoiceDate.Value)
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.Value = Trim(Me.txtPaymentDate.Value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtExpectPaymentMonth_AfterUpdate()
    Dim pos As Long
    Dim M As String

    Me.txtExpectPaymentMonth.Value = Trim(Me.txtExpectPaymentMonth.Value)

    If Me.txtExpectPaymentMonth = "" Then Exit Sub
    
    pos = InStr(Me.txtExpectPaymentMonth, "��")
    If pos <> 0 Then
        M = Left(Me.txtExpectPaymentMonth, pos - 1)
        Me.txtExpectPaymentDate.Value = DateSerial(Year(Date), M, 1)
        Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "��")
        Exit Sub
    End If
    
    If IsNumeric(Me.txtExpectPaymentMonth) Then
        Me.txtExpectPaymentDate.Value = DateSerial(Year(Date), Me.txtExpectPaymentMonth, 1)
        Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "��")
        Exit Sub
    End If
    
    Me.txtExpectPaymentDate.Value = Me.txtExpectPaymentMonth
    Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "��")
     
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub

