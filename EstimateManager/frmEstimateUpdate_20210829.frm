VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate_20210829 
   Caption         =   "���� ����"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17565
   OleObjectBlob   =   "frmEstimateUpdate_20210829.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateUpdate_20210829"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim orgManagementID As Variant


Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim db As Variant
    Dim contr As Control
    
    '������ �� ��ȣ
    cRow = Selection.row

    '�����Ͱ� �ִ� ���� �ƴ� ���� ����
    If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then End
    
    'Label ��ġ ���߱�
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '���� ������ �о����
    estimate = Get_Record_Array(shtEstimate, shtEstimateAdmin.Cells(cRow, 2))

    Me.txtID.Value = estimate(1)    'ID
    Me.txtEstimateName.Value = estimate(6)  '������
    Me.txtManagementID.Value = estimate(2)    '������ȣ
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
    InitializeCboProductonUnit  '��������׸� ����
    
    Me.txtExecutionCost.Value = Format(estimate(17), "#,##0")   '���డ
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
    Me.chkVAT.Value = estimate(31)      '�ΰ��� ���� ����
    
'    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '�Աݿ����
'    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '�Աݾ�
'    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '���Աݾ�
    
    Me.txtInsertDate.Value = estimate(23)    '�������
    Me.txtUpdateDate.Value = estimate(24)    '��������
    
    '���� �� ������ȣ
    orgManagementID = Me.txtManagementID
    
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtCustomer, True)

    Update_Cbo Me.cboCustomer, db
End Sub


Sub InitializeCboManager()
    Dim db As Variant
    Dim i As Long
    
    '����� DB�� �о�ͼ�
    db = Get_DB(shtManager, True)
    '�ŷ�ó������ ���͸�
    db = Filtered_DB(db, Me.cboCustomer.Value, 1)
    
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

Sub InitializeLstProduction()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2)
    
    'DB�� ���� ���� ���
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 10)) Then
                db(i, 10) = Format(db(i, 10), "#,##0")
            End If
            If IsNumeric(db(i, 11)) Then
                '��� �հ� ����
                totalCost = totalCost + CLng(db(i, 11))
                '���� ���� 1,000�ڸ� ó��
                db(i, 11) = Format(db(i, 11), "#,##0")
            End If

        Next
        
        Me.txtProductionTotalCost.Value = totalCost
        Me.txtProductionTotalCost.Text = Format(totalCost, "#,##0")
      
        Update_List Me.lstProductionList, db, "0pt;0pt;0pt,50pt,120pt;60pt;60pt;30pt;30pt;55pt;55pt;110pt;0pt"

    End If
    
    Me.txtProductionID.Value = ""
    
End Sub

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
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
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 3, orgManagementID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    '������ ������Ʈ
    Update_Record shtEstimate, Me.txtID.Value, _
        Me.txtManagementID.Value, Me.txtLinkedID.Value, _
        Me.cboCustomer.Value, Me.cboManager.Value, _
        Me.txtEstimateName.Value, Me.txtSize.Value, _
        Me.txtAmount.Value, Me.cboUnit.Value, _
        Me.txtUnitPrice.Value, Me.txtEstimatePrice.Value, _
        Me.txtEstimateDate.Value, Me.txtBidDate.Value, _
        Me.txtAcceptedDate.Value, Me.txtDeliveryDate.Value, _
        Me.txtInsuranceDate.Value, Me.txtExecutionCost.Value, _
        Me.txtBidPrice.Value, Me.txtBidMargin.Value, _
        Me.txtBidMarginRate.Value, Me.txtAcceptedPrice.Value, _
        Me.txtAcceptedMargin.Value, _
        Me.txtInsertDate.Value, Date, _
        Me.cboCategory.Value, Me.txtSpecificationDate.Value, _
        Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, _
        Me.txtExpectPaymentDate.Value, Me.txtVAT.Value, Me.chkVAT.Value

    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    
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

    '���װ� ������ ���
    If Me.txtBidPrice.Value <> "" And Me.txtExecutionCost.Value <> "" Then
        '���� = ������ - ���డ
        Me.txtBidMargin.Value = CLng(Me.txtBidPrice.Value) - CLng(Me.txtExecutionCost.Value)
        Me.txtBidMargin.Text = Format(Me.txtBidMargin.Value, "#,##0")
        '������ = ���� / ������
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


Sub InsertProduction()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "ǰ���� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "�ݾ��� �Է��ϼ���.": Exit Sub
    
    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    '��������׸� ����
    Insert_Record shtProduction, CLng(Me.txtID.Value), Me.txtManagementID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Date
    
    '��������׸� �հ� ���
    Me.txtProductionTotalCost.Value = GetProductionTotalCost
    Me.txtExecutionCost.Value = Me.txtProductionTotalCost.Value

    '���డ �������� ��� �ٽ� ���
    CalculateEstimateUpdateCost
    
    '������డ, ��������, ������, �������� �ݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "���డ", CLng(Me.txtProductionTotalCost.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "����", CLng(Me.txtBidMargin.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "������", Me.txtBidMarginRate.Value
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "��������", CLng(Me.txtAcceptedMargin.Value)
    
    Me.txtProductionID.Value = ""
    
    '��������׸� ����Ʈ�ڽ� ���ΰ�ħ
    InitializeLstProduction
    
End Sub


Sub UpdateProduction()
    Dim cost As Variant

    If Me.txtProductionID.Value = "" Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    If Me.txtProductionItem.Value = "" Then MsgBox "ǰ���� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "�ݾ��� �Է��ϼ���.": Exit Sub

    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    '���� ��������׸� ������Ʈ
    Update_Record shtProduction, Me.txtProductionID.Value, Me.txtID.Value, Me.txtManagementID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Date
    
    '������డ ���
    Me.txtProductionTotalCost.Value = GetProductionTotalCost
    Me.txtExecutionCost.Value = Me.txtProductionTotalCost.Value
    
    '������డ �������� ��� �ٽ� ���
    CalculateEstimateUpdateCost
    
    '������డ, ��������, ������, �������� �ݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "���డ", CLng(Me.txtProductionTotalCost.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "����", CLng(Me.txtBidMargin.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "������", Me.txtBidMarginRate.Value
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "��������", CLng(Me.txtAcceptedMargin.Value)
    
    InitializeLstProduction
    
    Select_ListItm Me.lstProductionList, Me.txtProductionID.Value
    
End Sub


Sub DeleteProduction()
    Dim db As Variant
    Dim YN As VbMsgBoxResult

    If Me.txtProductionID.Value = "" Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
        
    '�ȳ� ���� ���
    YN = MsgBox("������ �׸��� �����Ͻðڽ��ϱ�? ������ ������ ������ �Ұ����մϴ�.", vbYesNo)
    If YN = vbNo Then Exit Sub

    '��������׸񿡼� ����
    Delete_Record shtProduction, Me.txtProductionID.Value

    '������డ ���
    Me.txtProductionTotalCost.Value = GetProductionTotalCost
    Me.txtExecutionCost.Value = Me.txtProductionTotalCost.Value
    
     '������డ �������� ��� �ٽ� ���
     CalculateEstimateUpdateCost

     '������డ, ��������, ������, �������� �ݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "���డ", CLng(Me.txtProductionTotalCost.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "����", CLng(Me.txtBidMargin.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "������", Me.txtBidMarginRate.Value
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "��������", CLng(Me.txtAcceptedMargin.Value)

    Me.txtProductionID.Value = ""

    InitializeLstProduction

    ClearProductionInput
    
End Sub

Sub ProductionToOrder()

    If Me.txtProductionID.Value = "" Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    '������ ��������׸��� ���� ���̺� ���
    Insert_Record shtOrder, _
        , , Me.txtManagementID.Value, _
        Me.txtProductionCustomer.Value, _
        Me.txtProductionItem.Value, _
        Me.txtProductionMaterial.Value, _
        Me.txtProductionSize.Value, _
        Me.txtProductionAmount.Value, _
        Me.cboProductionUnit.Value, _
        Me.txtProductionUnitPrice.Value, _
        Me.txtProductionCost.Value, _
        , _
        , _
        , , , _
        , , , , , _
        Date, , Me.txtID, False, Me.txtProductionMemo
        
        shtOrderAdmin.Activate
        shtOrderAdmin.OrderSearch
        shtOrderAdmin.GoToEnd
        
        MsgBox "'" & Me.txtProductionItem & "' �׸��� �����Ͽ����ϴ�."
    End Sub

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2)
    
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

Sub ClearProductionInput()
    Me.txtProductionID.Value = ""
    Me.txtProductionCustomer.Value = ""
    Me.txtProductionItem.Value = ""
    Me.txtProductionMaterial.Value = ""
    Me.txtProductionSize.Value = ""
    Me.txtProductionAmount.Value = ""
    Me.cboProductionUnit.Value = ""
    Me.txtProductionUnitPrice.Value = ""
    Me.txtProductionCost.Value = ""
    Me.txtProductionMemo.Value = ""
End Sub

Private Sub btnEstimateUpdate_Click()
    UpdateEstimate
End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
    
    '�������� ȭ�� ���ΰ�ħ
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
End Sub

Private Sub lstProductionList_Click()
    Dim arr As Variant

    arr = Get_ListItm(Me.lstProductionList)
    Me.txtProductionID.Value = arr(0)                       'ID
    Me.txtProductionCustomer = arr(3)               '�ŷ�ó
    Me.txtProductionItem.Value = arr(4)                     'ǰ��
    Me.txtProductionMaterial.Value = arr(5)           '����
    Me.txtProductionSize.Value = arr(6)                '�԰�
    Me.txtProductionAmount.Value = arr(7)           '����
    Me.cboProductionUnit.Value = arr(8)               '����
    Me.txtProductionUnitPrice.Value = arr(9)        '�ܰ�
    Me.txtProductionUnitPrice.Text = Format(arr(9), "#,##0")
    Me.txtProductionCost.Value = arr(10)         '�ݾ�
    Me.txtProductionCost.Text = Format(arr(10), "#,##0")
    Me.txtProductionMemo = arr(11)       '�޸�
    
End Sub

Private Sub btnProductionClear_Click()
    ClearProductionInput
End Sub

Private Sub btnProductionDelete_Click()
    DeleteProduction
End Sub

Private Sub btnProductionInsert_Click()
    InsertProduction
End Sub

Private Sub btnProductionUpdate_Click()
    UpdateProduction
End Sub

Private Sub btnProductionToOrder_Click()
    ProductionToOrder
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
End Sub

Private Sub cboCustomer_Change()
    InitializeCboManager
End Sub

Private Sub txtEstimateDate_Change()
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.lblErrorMessage.Visible = False
End Sub

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

Private Sub txtProductionAmount_AfterUpdate()
    
    If Me.txtProductionAmount.Value = "" Then
        Exit Sub
    End If
    
    If IsNumeric(Me.txtProductionAmount.Value) Then
        Me.txtProductionAmount.Text = Format(Me.txtProductionAmount.Value, "#,##0")
        
        '�ݾ� = ���� * �ܰ�
        If IsNumeric(Me.txtProductionUnitPrice.Value) Then
            Me.txtProductionCost.Value = CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value)
            Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
        End If
    End If
End Sub

Private Sub txtProductionUnitPrice_AfterUpdate()
    
    If Me.txtProductionUnitPrice.Value = "" Then
        Exit Sub
    End If
    
    If IsNumeric(Me.txtProductionUnitPrice.Value) Then
        Me.txtProductionUnitPrice.Text = Format(Me.txtProductionUnitPrice.Value, "#,##0")
        
        If Me.txtProductionAmount.Value = "" Then
            Me.txtProductionCost.Value = Me.txtProductionUnitPrice.Value
            Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
        Else
            If IsNumeric(Me.txtProductionAmount.Value) Then
                '�ݾ� = ���� * �ܰ�
                Me.txtProductionCost.Value = CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value)
                Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
            End If
        End If
    End If
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub








