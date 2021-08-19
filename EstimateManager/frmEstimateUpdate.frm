VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "���� ���� ����"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17730
   OleObjectBlob   =   "frmEstimateUpdate.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim orgEstimateID As Variant


Private Sub btnEstimateClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub

Private Sub btnEstimateUpdate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UpdateEstimate
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

'�������� �Է� �ڽ�
Private Sub txtAcceptedDate_AfterUpdate()
    EveryCostUpdate
End Sub

'���Ǻ��� ���� �Է¹ڽ�
Private Sub txtInsuranceDate_AfterUpdate()
    EveryCostUpdate
End Sub

'�ŷ����� ���� �Է¹ڽ�
Private Sub txtSpecificationDate_AfterUpdate()
    EveryCostUpdate
End Sub

'���ݰ�꼭 ���� �Է¹ڽ�
Private Sub txtTaxInvoiceDate_AfterUpdate()
   EveryCostUpdate
End Sub

'�������� �Է¹ڽ�
Private Sub txtPaymentDate_AfterUpdate()
    EveryCostUpdate
End Sub

'����������� �Է¹ڽ�
Private Sub txtExpectPaymentDate_AfterUpdate()
    EveryCostUpdate
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
    
    '��� �ʵ� ���
    EveryCostUpdate
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
    
    '��� �ʵ� ���
    EveryCostUpdate
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
    
    '��� �ʵ� ���
    EveryCostUpdate
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

    '��� �ʵ� ���
    EveryCostUpdate
    
End Sub

'�������� Ķ���� ����
Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    EveryCostUpdate
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
    EveryCostUpdate
End Sub

'�������� Ķ���� ����
Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
    EveryCostUpdate
End Sub

'�ŷ����� Ķ���� ����
Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    EveryCostUpdate
End Sub

'���ݰ�꼭 Ķ���� ����
Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    EveryCostUpdate
End Sub

'����������� Ķ���� ����
Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    EveryCostUpdate
End Sub


Private Sub lstProductionList_Click()
    Dim arr As Variant

    arr = Get_ListItm(Me.lstProductionList)
    
    Me.txtProductionID = arr(0)
    Me.txtProductionItem = arr(2)
    If IsNumeric(arr(3)) Then Me.txtProductionCost = CLng(arr(3)) Else Me.txtProductionCost = arr(3)
    Me.txtProductionMemo = arr(4)
    
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim manager As Variant
    Dim customer As Variant
    Dim DB As Variant
    
    '������ �� ��ȣ
    cRow = Selection.row

    '�����Ͱ� �ִ� ���� �ƴ� ���� ����
    If cRow < 8 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then End
    
    '����/�����/�ŷ�ó ������ �о����
    estimate = Get_Record_Array(shtEstimate, shtEstimateAdmin.Cells(cRow, 2))
    manager = Get_Record_Array(shtManager, estimate(2))
    customer = Get_Record_Array(shtCustomer, manager(2))

    Me.txtID.Value = estimate(1)    'ID
    Me.txtManagerID.Value = estimate(2) 'ID_�����
    Me.txtEstimateName.Value = estimate(5)  '������
    Me.txtEstimateID.Value = estimate(3)    '������ȣ
    Me.txtLinkedID.Value = estimate(4)  '�����ȣ
    
    InitializeCboCustomer
    Select_CboItm Me.cboCustomer, customer(1), 1    '�ŷ�ó
    InitializeCboManager
    Select_CboItm Me.cboManager, manager(1), 1  '�����
    
    Me.txtSize.Value = estimate(6)  '�԰�
    
    InitializeCboUnit
    Me.cboUnit.Value = estimate(8)  '����, ID�� �����Ƿ� ���� value ������ ���õ�
    
    Me.txtAmount.Value = estimate(7)    '����
    Me.txtUnitPrice.Value = estimate(9)     '�����ܰ�
    Me.txtEstimatePrice.Value = estimate(10)    '�����ݾ�
    
    Me.txtEstimateDate.Value = estimate(11)    '��������
    Me.txtBidDate.Value = estimate(12)    '��������
    Me.txtAcceptedDate.Value = estimate(13)    '��������
    Me.txtDeliveryDate.Value = estimate(14)    '��ǰ����
    Me.txtInsuranceDate.Value = estimate(15)    '��������
    
    InitializeLstProduction    '������� �Է¸��
    Me.txtProductionTotalCost.Value = estimate(16)    '������డ
    
    Me.txtBidPrice.Value = estimate(17)    '������
    Me.txtBidMargin.Value = estimate(18)    '����
    Me.txtBidMarginRate.Value = estimate(19)    '������
    Me.txtAcceptedPrice.Value = estimate(20)    '���ֱݾ�
    Me.txtAcceptedMargin.Value = estimate(21)    '��������
    
    Me.txtSpecificationDate.Value = estimate(22)    '�ŷ�����
    Me.txtTaxInvoiceDate.Value = estimate(23)    '���ݰ�꼭
    Me.txtPaymentDate.Value = estimate(24)    '��������
    Me.txtExpectPaymentDate.Value = estimate(25)    '�����������
    Me.txtVAT.Value = estimate(26)    '�ΰ���
    Me.txtExpectPay.Value = estimate(27)    '�Աݿ����
    Me.txtPaid.Value = estimate(28)    '�Աݾ�
    Me.txtUnpaid.Value = estimate(29)    '���Աݾ�
    
    Me.txtInsertDate.Value = estimate(30)    '�������
    Me.txtUpdateDate.Value = estimate(31)    '��������
    
    '���� �� ������ȣ
    orgEstimateID = Me.txtEstimateID
    
End Sub


Sub UpdateEstimate()
    Dim DB As Variant
    Dim blnUnique As Boolean
    
    '�Է� ������ üũ
    If InputValidationCheck = False Then
        Exit Sub
    End If

    '�������� DB �о����
    DB = Get_DB(shtEstimate)
    
    '������ ������ȣ�� �ִ��� üũ
    blnUnique = IsUnique(DB, Me.txtEstimateID.Value, 3, orgEstimateID)
    If blnUnique = False Then MsgBox "������ ������ȣ�� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbExclamation: Exit Sub
    
    '�����ݾ� ��� = �����ܰ� * ����
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value * Me.txtAmount.Value
    End If
    
    '������ ������Ʈ
    Update_Record shtEstimate, Me.txtID.Value, Me.cboManager.Value, _
        Me.txtEstimateID.Value, Me.txtLinkedID.Value, _
        Me.txtEstimateName.Value, Me.txtSize.Value, _
        Me.txtAmount.Value, Me.cboUnit.Value, _
        Me.txtUnitPrice.Value, Me.txtEstimatePrice.Value, _
        Me.txtEstimateDate.Value, Me.txtBidDate.Value, _
        Me.txtAcceptedDate.Value, Me.txtDeliveryDate.Value, _
        Me.txtInsuranceDate.Value, Me.txtProductionTotalCost.Value, _
        Me.txtBidPrice.Value, Me.txtBidMargin.Value, _
        Me.txtBidMarginRate.Value, Me.txtAcceptedPrice.Value, _
        Me.txtAcceptedMargin.Value, Me.txtSpecificationDate.Value, _
        Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, _
        Me.txtExpectPaymentDate.Value, Me.txtVAT.Value, _
        Me.txtExpectPay.Value, Me.txtPaid.Value, _
        Me.txtUnpaid.Value, _
        Me.txtInsertDate.Value, Date

    Unload Me
    
    shtEstimateAdmin.EstimateSearch
    
End Sub

Sub InitializeCboCustomer()
    Dim DB As Variant
    DB = Get_DB(shtCustomer)

    Update_Cbo Me.cboCustomer, DB, 2
End Sub

Sub InitializeCboManager()
    Dim DB As Variant
    Dim i As Long
    
    '����� DB�� �о�ͼ�
    DB = Get_DB(shtManager)
    '�ŷ�óID�� ���͸�
    DB = Filtered_DB(DB, Me.cboCustomer.Value, 2)
    
    '���� �޺��ڽ� ���������
    Me.cboManager.Clear
    
    '����ڰ� ������ �޺��ڽ��� �߰���
    If Not IsEmpty(DB) Then
        'Filtered_DB ����ϸ鼭 ID�� ���ڿ��� �ٲ� -> �̰� ���ڷ� ��ȯ
        For i = 1 To UBound(DB, 1)
            DB(i, 1) = Val(DB(i, 1))
            DB(i, 2) = Val(DB(i, 2))
        Next
        
        Update_Cbo Me.cboManager, DB, 3
    End If
End Sub

Sub InitializeCboUnit()
    Dim DB As Variant
    DB = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, DB
End Sub

Sub InitializeLstProduction()
    Dim DB As Variant
    Dim i, totalCost As Long
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    DB = Get_DB(shtProduction)
    DB = Filtered_DB(DB, Me.txtID.Value, 2)
    
    Update_List Me.lstProductionList, DB, "0pt;0pt;60pt;50pt;100pt;"
    
    '��� �հ� ����
    If Not IsEmpty(DB) Then
        For i = 1 To UBound(DB)
            If IsNumeric(DB(i, 4)) Then
                totalCost = totalCost + CLng(DB(i, 4))
            End If
        Next
    End If
    Me.txtProductionSum = totalCost
    
End Sub

Sub InitalizeProductionInput()
    Me.txtProductionID.Value = ""
    Me.txtProductionItem.Value = ""
    Me.txtProductionCost.Value = ""
    Me.txtProductionMemo.Value = ""
End Sub

Sub InsertProjection()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "�׸��� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "����� �Է��ϼ���.": Exit Sub
    
    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    Insert_Record shtProduction, CLng(Me.txtID.Value), Me.txtProductionItem.Value, cost, Me.txtProductionMemo.Value
    
    InitializeLstProduction
    
    InitalizeProductionInput
    
End Sub


Sub UpdateProjection()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "�׸��� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "����� �Է��ϼ���.": Exit Sub

    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
        
    Update_Record shtProduction, Me.txtProductionID.Value, Me.txtID.Value, Me.txtProductionItem.Value, cost, Me.txtProductionMemo.Value

    InitializeLstProduction
    
    Select_ListItm Me.lstProductionList, Me.txtProductionID.Value

End Sub


Sub DeleteProjection()
    Dim DB As Variant
    Dim YN As VbMsgBoxResult

    Delete_Record shtProduction, Me.txtProductionID.Value

    InitializeLstProduction
    
    InitalizeProductionInput
End Sub

Function InputValidationCheck()
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
    
    '������ �ԷµǾ����� üũ
    If Me.txtAmount.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "������ �Է��ϼ���."
    End If
    
    '�����ܰ��� �ԷµǾ����� üũ
    If Me.txtUnitPrice.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "�����ܰ��� �Է��ϼ���."
    End If
    
    '�������ڰ� �ԷµǾ����� üũ
    If Me.txtEstimateDate.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "�������ڸ� �Է��ϼ���."
    End If
    
    If bCorrect = False Then
        Me.lblErrorMessage.Visible = True
    Else
        Me.lblErrorMessage.Visible = False
    End If
    
    InputValidationCheck = bCorrect
End Function

Sub EveryCostUpdate()

    '�����ݾ� ���
    '�������� �����̸� �����ݾ��� �����ܰ�
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
    End If

    '���װ� ������ ���
    If Me.txtBidPrice.Value <> "" And Me.txtProductionTotalCost <> "" Then
        '���� = �����ݾ� - �������ݾ�
        Me.txtBidMargin.Value = CLng(Me.txtBidPrice.Value) - CLng(Me.txtProductionTotalCost.Value)
        '������ = ���� / �����ݾ�
        Me.txtBidMarginRate.Value = CLng(Me.txtBidMargin.Value) / CLng(Me.txtBidPrice.Value)
    Else
        Me.txtBidMargin.Value = 0
    End If

    '���ֱݾ� ���
    '�������ڰ� �ִ� ��츸
    If Me.txtAcceptedDate.Value = "" Then
        Me.txtAcceptedPrice.Value = 0
        Me.txtAcceptedMargin.Value = 0
    Else
        '���ֱݾ��� �����ݾ����� ����
        Me.txtAcceptedPrice.Value = Me.txtBidPrice
        '���������� �������� ����
        Me.txtAcceptedMargin.Value = Me.txtBidMargin.Value
    End If

    '�ΰ��� ���
    '���ݰ�꼭 ���ڰ� �ִ� ��츸
    If Me.txtTaxInvoiceDate.Value = "" Then
        Me.txtVAT.Value = 0
    Else
        '�ΰ����� ���ֱݾ��� 10%
        If Me.txtAcceptedPrice.Value <> "" And Me.txtAcceptedPrice.Value <> 0 Then
            Me.txtVAT.Value = CLng(Me.txtAcceptedPrice.Value) * 0.1
        End If
    End If

    '�Աݿ���� ���
    If Me.txtTaxInvoiceDate.Value = "" Then
        '���ݰ�꼭 ���ڰ� ���� ���� ���ֱݾ�
        Me.txtExpectPay.Value = Me.txtAcceptedPrice
    Else
        '���ݰ�꼭 ���ڰ� �ִ� ���� ���ֱݾ�+�ΰ���
        Me.txtExpectPay.Value = CLng(Me.txtAcceptedPrice) + CLng(Me.txtVAT.Value)
    End If

    '�Աݾ� ���
    If Me.txtPaymentDate.Value = "" Then
        Me.txtPaid.Value = 0
    Else
        Me.txtPaid.Value = Me.txtExpectPay.Value
    End If
    
    '���Աݾ� ���
    Me.txtUnpaid.Value = CLng(Me.txtExpectPay.Value) - CLng(Me.txtPaid.Value)
    
End Sub

'=============================================
'����Ʈ�ڽ� ��ũ��
'Private Sub lstProductionList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstProductionList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstProductionList
'End Sub


