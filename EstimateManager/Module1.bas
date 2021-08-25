Attribute VB_Name = "Module1"
Option Explicit

Dim orgEstimateID As Variant


Private Sub btnEstimateClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub

Private Sub btnEstimateUpdate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UpdateEstimate
End Sub


Private Sub btnProductionClear_Change()

End Sub

Private Sub btnProductionClear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    InitalizeProductionInput
End Sub

Private Sub btnProductionDelete_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    DeleteProjection
End Sub

Private Sub btnProductionInsert_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    InsertProjection
End Sub

Private Sub btnProductionUpdate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UpdateProjection
End Sub


Private Sub cboCustomer_Change()
    '�޺��ڽ����� �ŷ�ó�� �����ϸ� �ش� �ŷ�ó�� ����ڷ� ����� �޺��ڽ��� ����
    InitializeCboManager
End Sub

'�ΰ��� ���� üũ
Private Sub chkVAT_Click()
    CalculateEstimateUpdateCost
End Sub


'�������� �Է� �ڽ�
Private Sub txtAcceptedDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'���Ǻ��� ���� �Է¹ڽ�
Private Sub txtInsuranceDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'�ŷ����� ���� �Է¹ڽ�
Private Sub txtSpecificationDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'���ݰ�꼭 ���� �Է¹ڽ�
Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateEstimateUpdateCost
End Sub

'�������� �Է¹ڽ�
Private Sub txtPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'����������� �Է¹ڽ�
Private Sub txtExpectPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtEstimateDate_Change()
    '���� �޽��� ����
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateID_AfterUpdate()
    '���� �޽��� ����
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    '���� �޽��� ����
    Me.lblErrorMessage.Visible = False
End Sub

'���� �Է�
Private Sub txtAmount_AfterUpdate()
    '�����޽��� ����
    Me.lblErrorMessage.Visible = False
    
    '�������� �����̸� ����
    If Me.txtAmount.Value <> "" Then
        '�������� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            Me.lblErrorMessage.Caption = "���ڸ� �Է��ϼ���."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '���� 1,000�ڸ� �ĸ� ó��
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    '��� �ʵ� ���
    CalculateEstimateUpdateCost
End Sub

'�����ܰ� �Է�
Private Sub txtUnitPrice_AfterUpdate()
     '�����޽��� ����
    Me.lblErrorMessage.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '�����ܰ����� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            Me.lblErrorMessage.Caption = "���ڸ� �Է��ϼ���."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '�����ܰ� 1,000�ڸ� �ĸ� ó��
    Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    
    '��� �ʵ� ���
    CalculateEstimateUpdateCost
End Sub

'���� ���� �ùķ��̼� ��� �Է�
Private Sub txtProductionCost_AfterUpdate()
    '�����޽��� ����
    Me.lblErrorMessage.Visible = False
    
    If Me.txtProductionCost.Value = "" Then
        Exit Sub
    End If
    
    '��� �Է°��� ���ڰ� �ƴ� ��� �����޽��� ���
    If Not IsNumeric(Me.txtProductionCost.Value) Then
        Me.txtProductionCost.Value = ""
        Me.lblErrorMessage.Caption = "���ڸ� �Է��ϼ���."
        Me.lblErrorMessage.Visible = True
        Exit Sub
    End If
    
    '�հ� �ݾ� 1,000�ڸ� �ĸ� ó��
    Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
End Sub

'������డ �Է�
Private Sub txtProductionTotalCost_AfterUpdate()
     '�����޽��� ����
    Me.lblErrorMessage.Visible = False
    
    If Me.txtProductionTotalCost.Value <> "" Then
        '������డ ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtProductionTotalCost.Value) Then
            Me.txtProductionTotalCost.Value = ""
            Me.lblErrorMessage.Caption = "���ڸ� �Է��ϼ���."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '���� ���� �ݾ� 1,000�ڸ� �ĸ� ó��
    Me.txtProductionTotalCost.Text = Format(Me.txtProductionTotalCost.Value, "#,##0")
    
    '��� �ʵ� ���
    CalculateEstimateUpdateCost
End Sub

'�����ݾ� �Է�
Private Sub txtBidPrice_AfterUpdate()
     '�����޽��� ����
    Me.lblErrorMessage.Visible = False
    
    If Me.txtBidPrice.Value <> "" Then
        '�����ݾ��� ���ڰ� �ƴ� ��� �����޽��� ���
        If Not IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtBidPrice.Value = ""
            Me.lblErrorMessage.Caption = "���ڸ� �Է��ϼ���."
            Me.lblErrorMessage.Visible = True
        End If
    End If

    '�����ݾ� 1,000�ڸ� �ĸ� ó��
    Me.txtBidPrice.Text = Format(Me.txtBidPrice.Value, "#,##0")
    
    '��� �ʵ� ���
    CalculateEstimateUpdateCost
    
End Sub

'�������� Ķ���� ����
Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

'�������� Ķ���� ����
Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

'��ǰ���� Ķ���� ����
Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

'�������� Ķ���� ����
Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

'���Ǻ��� Ķ���� ����
Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

'�������� Ķ���� ����
Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
    CalculateEstimateUpdateCost
End Sub

'�ŷ����� Ķ���� ����
Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

'���ݰ�꼭 Ķ���� ����
Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateEstimateUpdateCost
End Sub

'����������� Ķ���� ����
Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    CalculateEstimateUpdateCost
End Sub


Private Sub lstProductionList_Click()
    Dim arr As Variant

    arr = Get_ListItm(Me.lstProductionList)
    Me.txtProductionID.Value = arr(0)                       'ID
    Select_CboItm Me.cboCustomer2, Trim(arr(3)), 1    '�ŷ�ó
    Me.txtProductionItem.Value = arr(4)                     'ǰ��
    Me.txtProductionAmount.Value = arr(5)           '����
    Me.txtProductionUnitPrice.Value = arr(6)        '�ܰ�
    Me.txtProductionUnitPrice.Text = Format(arr(6), "#,##0")
    Me.txtProductionCost.Value = arr(7)         '�ݾ�
    Me.txtProductionCost.Text = Format(arr(7), "#,##0")
    Me.txtProductionMemo = arr(8)       '�޸�
    
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim manager As Variant
    Dim customer As Variant
    Dim db As Variant
    
    '������ �� ��ȣ
    cRow = Selection.row

    '�����Ͱ� �ִ� ���� �ƴ� ���� ����
    If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then End
    
    '����/�����/�ŷ�ó ������ �о����
    estimate = Get_Record_Array(shtEstimate, shtEstimateAdmin.Cells(cRow, 2))

    Me.txtID.Value = estimate(1)    'ID
    Me.txtEstimateName.Value = estimate(6)  '������
    Me.txtEstimateID.Value = estimate(2)    '������ȣ
    Me.txtLinkedID.Value = estimate(3)  '�����ȣ
    
    InitializeCboCustomer
    Select_CboItm Me.cboCustomer, Trim(estimate(4)), 1    '�ŷ�ó
    InitializeCboManager
    Select_CboItm Me.cboManager, Trim(estimate(5)), 1  '�����
    
    Me.txtSize.Value = estimate(7)  '�԰�
    
    InitializeCboUnit
    Me.cboUnit.Value = Trim(estimate(9))  '����, ID�� �����Ƿ� ���� value ������ ���õ�
    
    Me.txtAmount.Value = Format(estimate(8), "#,##0")   '����
    Me.txtUnitPrice.Value = Format(estimate(10), "#,##0")     '�����ܰ�
    Me.txtEstimatePrice.Value = Format(estimate(11), "#,##0")     '�����ݾ�
    
    Me.txtEstimateDate.Value = estimate(12)    '��������
    Me.txtBidDate.Value = estimate(13)    '��������
    Me.txtAcceptedDate.Value = estimate(14)    '��������
    Me.txtDeliveryDate.Value = estimate(15)    '��ǰ����
    Me.txtInsuranceDate.Value = estimate(16)    '��������
    
    InitializeLstProduction    '��������׸� ���
    InitializeCboCustomer2   '��������׸� �ŷ�ó
    Me.txtProductionTotalCost.Value = Format(estimate(17), "#,##0")    '������డ
    Me.txtProductionID.Value = ""
    
    Me.txtBidPrice.Value = Format(estimate(18), "#,##0")    '������
    Me.txtBidMargin.Value = Format(estimate(19), "#,##0")    '����
    Me.txtBidMarginRate.Value = Format(estimate(20), "0.0%")    '������
    Me.txtAcceptedPrice.Value = Format(estimate(21), "#,##0")    '���ֱݾ�
    Me.txtAcceptedMargin.Value = Format(estimate(22), "#,##0")   '��������
    
    InitializeCboCategory
    Me.cboCategory.Value = Trim(estimate(25))   '�з�
    Me.txtSpecificationDate.Value = estimate(26)    '�ŷ�����
    Me.txtTaxInvoiceDate.Value = estimate(27)    '���ݰ�꼭
    Me.txtPaymentDate.Value = estimate(28)    '��������
    Me.txtExpectPaymentDate.Value = estimate(29)    '�����������
    Me.txtVAT.Value = Format(estimate(30), "#,##0")    '�ΰ���
    Me.chkVAT.Value = estimate(31)
'    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '�Աݿ����
'    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '�Աݾ�
'    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '���Աݾ�
    
    Me.txtInsertDate.Value = estimate(23)    '�������
    Me.txtUpdateDate.Value = estimate(24)    '��������
    
    
    '���� �� ������ȣ
    orgEstimateID = Me.txtEstimateID
    
    
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
    blnUnique = IsUnique(db, Me.txtEstimateID.Value, 3, orgEstimateID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    '�����ݾ� ��� = �����ܰ� * ����
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value * Me.txtAmount.Value
    End If
    
    '������ ������Ʈ
    Update_Record shtEstimate, Me.txtID.Value, _
        Me.txtEstimateID.Value, Me.txtLinkedID.Value, _
        Me.cboCustomer.Value, Me.cboManager.Value, _
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
        Me.cboCategory.Value, Me.txtSpecificationDate.Value, _
        Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, _
        Me.txtExpectPaymentDate.Value, Me.txtVAT.Value, Me.chkVAT.Value

    Unload Me
    
    shtEstimateAdmin.EstimateSearch
    
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtCustomer, True)

    Update_Cbo Me.cboCustomer, db
End Sub

Sub InitializeCboCustomer2()
    Dim db As Variant
    db = Get_DB(shtCustomer, True)

    Update_Cbo Me.cboCustomer2, db
End Sub

Sub InitializeCboManager()
    Dim db As Variant
    Dim i As Long
    
    '����� DB�� �о�ͼ�
    db = Get_DB(shtManager, True)
    '�ŷ�ó������ ���͸�
    db = Filtered_DB(db, Me.cboCustomer.Value, 2)
    
    '���� �޺��ڽ� ���������
    Me.cboManager.Clear
    
    '����ڰ� ������ �޺��ڽ��� �߰���
    If Not IsEmpty(db) Then
        Update_Cbo Me.cboManager, db, 1
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

Sub InitializeLstProduction()
    Dim db As Variant
    Dim i, totalCost As Long
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2)
    
    'DB�� ���� ���� ���
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 8)) Then
                '��� �հ� ����
                totalCost = totalCost + CLng(db(i, 8))
                '���� ���� 1,000�ڸ� ó��
                db(i, 8) = Format(db(i, 8), "#,##0")
            End If
        Next
        
        Me.txtProductionTotalCost = Format(totalCost, "#,##0")
        
        Update_List Me.lstProductionList, db, "0pt;0pt;0pt,50pt,130pt;20pt;50pt;50pt;130pt;0pt"
        
    End If
    
End Sub

Sub InitalizeProductionInput()
    Me.txtProductionID.Value = ""
    Me.cboCustomer2.Value = ""
    Me.txtProductionItem.Value = ""
    Me.txtProductionAmount.Value = ""
    Me.txtProductionUnitPrice.Value = ""
    Me.txtProductionCost.Value = ""
    Me.txtProductionMemo.Value = ""
End Sub

Sub InsertProjection()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "ǰ���� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "�ݾ��� �Է��ϼ���.": Exit Sub
    
    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    Insert_Record shtProduction, CLng(Me.txtID.Value), Me.txtEstimateID.Value, Me.cboCustomer2.Value, Me.txtProductionItem.Value, _
            Me.txtProductionAmount, Me.txtProductionUnitPrice, Me.txtProductionCost, Me.txtProductionMemo.Value, Date
    
    Me.txtProductionID.Value = ""
    
    InitializeLstProduction
    InitalizeProductionInput
    
End Sub


Sub UpdateProjection()
    Dim cost As Variant

    If Me.txtProductionID.Value = "" Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    If Me.txtProductionItem.Value = "" Then MsgBox "ǰ���� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "�ݾ��� �Է��ϼ���.": Exit Sub

    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
        
    Update_Record shtProduction, Me.txtProductionID.Value, Me.txtID.Value, Me.txtEstimateID.Value, Me.cboCustomer2.Value, Me.txtProductionItem.Value, _
            Me.txtProductionAmount, Me.txtProductionUnitPrice, Me.txtProductionCost, Me.txtProductionMemo.Value, Date
    
    InitializeLstProduction
    
    Select_ListItm Me.lstProductionList, Me.txtProductionID.Value

End Sub


Sub DeleteProjection()
    Dim db As Variant
    Dim YN As VbMsgBoxResult

    If Me.txtProductionID.Value = "" Then
        MsgBox "������ �׸��� �����ϼ���."
        Exit Sub
    Else
        Delete_Record shtProduction, Me.txtProductionID.Value

        Me.txtProductionID.Value = ""
    
        InitializeLstProduction
    
        InitalizeProductionInput
    End If
    
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
    If Me.txtEstimateID.Value = "" Then
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
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")

    '���װ� ������ ���
    If Me.txtBidPrice.Value <> "" And Me.txtProductionTotalCost <> "" Then
        '���� = �����ݾ� - �������ݾ�
        Me.txtBidMargin.Value = CLng(Me.txtBidPrice.Value) - CLng(Me.txtProductionTotalCost.Value)
        Me.txtBidMargin.Text = Format(Me.txtBidMargin.Value, "#,##0")
        '������ = ���� / �����ݾ�
        Me.txtBidMarginRate.Value = CLng(Me.txtBidMargin.Value) / CLng(Me.txtBidPrice.Value)
        Me.txtBidMarginRate.Text = Format(Me.txtBidMarginRate.Value, "0.0%")
    Else
        Me.txtBidMargin.Value = 0
    End If

    '���ֱݾ� ���
    If Me.txtAcceptedDate.Value = "" Then
        '�������ڰ� ���� ���
        Me.txtAcceptedPrice.Value = 0
        Me.txtAcceptedMargin.Value = 0
    Else
        '�������ڰ� �ִ� ���
        '���ֱݾ��� �����ݾ����� ����
        If IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtAcceptedPrice.Value = CLng(Me.txtBidPrice.Value)
        Else
            Me.txtAcceptedPrice.Value = 0
        End If
        Me.txtAcceptedPrice.Text = Format(Me.txtAcceptedPrice.Value, "#,##0")
        
        '���������� �������� ����
        If IsNumeric(Me.txtBidMargin.Value) Then
            Me.txtAcceptedMargin.Value = CLng(Me.txtBidMargin.Value)
        Else
            Me.txtAcceptedMargin.Value = 0
        End If
        Me.txtAcceptedMargin.Text = Format(Me.txtAcceptedMargin.Value, "#,##0")
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

    '�Աݿ���� ���
    If Me.txtTaxInvoiceDate.Value = "" Then
        '���ݰ�꼭 ���ڰ� ���� ���� ���ֱݾ�
        Me.txtExpectPay.Value = Me.txtAcceptedPrice
    Else
        '���ݰ�꼭 ���ڰ� �ִ� ���� ���ֱݾ�+�ΰ���
        If Me.txtAcceptedPrice.Value <> "" Then
            Me.txtExpectPay.Value = CLng(Me.txtAcceptedPrice.Value) + CLng(Me.txtVAT.Value)
        End If
    End If
    Me.txtExpectPay.Text = Format(Me.txtExpectPay.Value, "#,##0")

    '�Աݾ� ���
    If Me.txtPaymentDate.Value = "" Then
        Me.txtPaid.Value = 0
    Else
        Me.txtPaid.Value = Me.txtExpectPay.Value
        Me.txtPaid.Text = Format(Me.txtPaid.Value, "#,##0")
    End If
    
    '���Աݾ� ���
    Me.txtUnpaid.Value = CLng(Me.txtExpectPay.Value) - CLng(Me.txtPaid.Value)
    Me.txtUnpaid.Text = Format(Me.txtUnpaid.Value, "#,##0")
    
End Sub

'=============================================
'����Ʈ�ڽ� ��ũ��
'Private Sub lstProductionList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstProductionList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstProductionList
'End Sub




