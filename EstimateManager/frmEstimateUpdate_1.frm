VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate_1 
   Caption         =   "���� ����"
   ClientHeight    =   13440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19125
   OleObjectBlob   =   "frmEstimateUpdate_1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmEstimateUpdate_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim orgManagementID As Variant
Dim orgExecutionCost As String
Dim totlalCheckCount As Long




Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim estimateId As Variant
    Dim db As Variant
    Dim contr As Control
    
    If clickEstimateId <> "" Then              '���ְ������� ����Ŭ���� ���
        If IsNumeric(clickEstimateId) Then
            estimateId = CLng(clickEstimateId)
        Else
            estimateId = clickEstimateId
        End If
        clickEstimateId = ""
    Else
        '������ �� ��ȣ
        cRow = Selection.row
    
        '�����Ͱ� �ִ� ���� �ƴ� ���� ����
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then
            MsgBox "������ ���� ���� ���� ������ �� �������� ��ư�� Ŭ���ϼ���."
            End
        End If
        
        estimateId = shtEstimateAdmin.Cells(cRow, 2)
    End If
    
    '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
    If contr.Name Like "lbl*" Then
        'contr.top = contr.top + 2
        contr.BackColor = RGB(242, 242, 242)
    End If
    Next
    
    '�� ��ġ ����
    If estimateUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateUpdateFormX
        Me.top = estimateUpdateFormY
    End If

    
    '���� ������ �о����
    estimate = Get_Record_Array(shtEstimate, estimateId)

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
    
    Me.txtExecutionCost.Value = Format(estimate(17), "#,##0")   '���డ
    orgExecutionCost = Me.txtExecutionCost.Value
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
    
'    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '�Աݿ����
'    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '�Աݾ�
'    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '���Աݾ�
    
    '���� �� ������ȣ
    orgManagementID = Me.txtManagementID
    
'    InitializeLswProductionList    '��������׸� ���
'    InitializeCboProductonUnit  '��������׸� ����
'    InitializeLswOrderCustomerAutoComplete   '���ְŷ�ó �ڵ��ϼ�
    
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

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2, True)
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�ŷ�ó", 50
        .ColumnHeaders.Add , , "����", 60
        .ColumnHeaders.Add , , "�԰�", 60
        .ColumnHeaders.Add , , "����", 30, lvwColumnRight
        .ColumnHeaders.Add , , "����", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 60, lvwColumnRight
        .ColumnHeaders.Add , , "�޸�", 110
        .ColumnHeaders.Add , , "�������", 0
        
        .CheckBoxes = True
        .ColumnHeaders(1).Position = 5
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 11)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 11))
                End If
                
                Set li = .ListItems.Add(, , db(i, 5))
                li.ListSubItems.Add , , db(i, 1)
                li.ListSubItems.Add , , db(i, 2)
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , db(i, 8)
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                li.ListSubItems.Add , , db(i, 13)
            Next
            
            Me.txtProductionTotalCost.Value = totalCost
            Me.txtProductionTotalCost.Text = Format(totalCost, "#,##0")
        End If
    End With
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

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
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
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�ŷ�ó", 50
        .ColumnHeaders.Add , , "����", 60
        .ColumnHeaders.Add , , "�԰�", 60
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
        .ColumnHeaders.Add , , "��������", 59, lvwColumnCenter
        
        .ColumnHeaders(1).Position = 5
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 11)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 11))
                End If
                
                Set li = .ListItems.Add(, , db(i, 7))   'ǰ��
                li.ListSubItems.Add , , db(i, 1)        'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_����
                li.ListSubItems.Add , , db(i, 5)        '������ȣ
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
                li.ListSubItems.Add , , db(i, 24)      '��������
            Next
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
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 3, orgManagementID)
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
        Me.txtInsuranceDate.Value, Me.txtExecutionCost.Value, _
        Me.txtBidPrice.Value, Me.txtBidMargin.Value, _
        Me.txtBidMarginRate.Value, Me.txtAcceptedPrice.Value, _
        Me.txtAcceptedMargin.Value, _
        Me.txtInsertDate.Value, Date, _
        Me.cboCategory.Value, , _
        Me.txtSpecificationDate.Value, Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, Me.txtExpectPaymentDate.Value, _
        Me.txtVAT.Value, Me.txtMemo.Value, Me.chkVAT.Value
    
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
    
    '��������׸� ����Ʈ�ڽ� ���ΰ�ħ
    InitializeLswProductionList
    
    '����� ������ ����
    Me.txtProductionID.Value = Get_LastID(shtProduction)
    SelectItemLswProduction Me.txtProductionID.Value
    
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
    
    InitializeLswProductionList
    SelectItemLswProduction Me.txtProductionID.Value
    
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

    InitializeLswProductionList

    ClearProductionInput
    
End Sub

Sub ProductionToOrder()
    Dim li As ListItem
    Dim count As Long
    Dim managementID, customer, Item, material, size, amount, unit, unitPrice, cost, memo As Variant
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Checked = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "������ �׸��� üũ�ڽ��� üũ�ϼ���.": Exit Sub
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Checked = True Then
            Item = li.Text
            managementID = li.SubItems(3)
            customer = li.SubItems(4)
            material = li.SubItems(5)
            size = li.SubItems(6)
            amount = li.SubItems(7)
            unit = li.SubItems(8)
            unitPrice = li.SubItems(9)
            cost = li.SubItems(10)
            memo = li.SubItems(11)
            
            '������ ��������׸��� ���� ���̺� ���
            Insert_Record shtOrder, _
                , , managementID, customer, Item, material, size, amount, unit, unitPrice, cost, _
                , , , , , _
                , , , , , _
                Date, , Me.txtID, False, memo
                
            count = count + 1
        End If
    Next
    
    InitializeLswOrderList
    
    MsgBox "�� " & count & "�� �׸��� �����Ͽ����ϴ�.", vbInformation
    
    shtOrderAdmin.Activate
    shtOrderAdmin.OrderSearch
    shtOrderAdmin.GoToEnd

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

Sub SelectItemLswProduction(selectedID As Variant)
    Dim i As Long
    
    With Me.lswProductionList
        If Not IsMissing(selectedID) Then
            For i = 1 To .ListItems.count
                If selectedID = .ListItems(i).SubItems(1) Then
                    .SelectedItem = .ListItems(i)
                    .SetFocus
                End If
            Next
        End If
    End With
End Sub

Private Sub lswProductionList_Click()
    With Me.lswProductionList
        If Not .SelectedItem Is Nothing Then
            Me.txtProductionID.Value = .SelectedItem.ListSubItems(1)
            Me.txtProductionItem.Value = .SelectedItem.Text
            Me.txtProductionCustomer.Value = .SelectedItem.ListSubItems(4)
            Me.txtProductionMaterial.Value = .SelectedItem.ListSubItems(5)
            Me.txtProductionSize.Value = .SelectedItem.ListSubItems(6)
            Me.txtProductionAmount.Value = .SelectedItem.ListSubItems(7)
            Me.cboProductionUnit.Value = .SelectedItem.ListSubItems(8)
            Me.txtProductionUnitPrice.Value = .SelectedItem.ListSubItems(9)
            Me.txtProductionCost.Value = .SelectedItem.ListSubItems(10)
            Me.txtProductionMemo.Value = .SelectedItem.ListSubItems(11)
        End If
    End With
End Sub

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
`    UpdateEstimate
    
    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.Range("H" & selectionRow).Select
End Sub

Private Sub btnEstimateClose_Click()
    If orgExecutionCost <> Me.txtExecutionCost.Value Then
        Unload Me
        
        '���డ�� ����� ��쿡�� �������� ȭ�� ���ΰ�ħ
        shtEstimateAdmin.Activate
        shtEstimateAdmin.EstimateSearch
    Else
        Unload Me
    End If
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

Private Sub lswProductionList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    With Me.lswProductionList
        .SelectedItem.Selected = False
        
        If Item.Checked = True Then
            Item.Bold = True
            Item.ForeColor = vbBlue
            totlalCheckCount = totlalCheckCount + 1
        Else
            Item.Bold = False
            Item.ForeColor = vbBlack
            totlalCheckCount = totlalCheckCount - 1
        End If
    End With
    
    If totlalCheckCount = 0 Then
        Me.btnProductionToOrder.Caption = "üũ �׸� ����"
    Else
        Me.btnProductionToOrder.Caption = totlalCheckCount & "�� �׸� ����"
    End If
End Sub


Private Sub txtProductionCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '����Ű - ���� �Է�ĭ���� �̵�
        Me.lswOrderCustomerAutoComplete.Visible = False
        Me.txtProductionItem.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '��Ű, �Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtProductionCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '�ŷ�ó �ڵ��ϼ� ó��
    With Me.lswOrderCustomerAutoComplete
        If Me.txtProductionCustomer.Value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '���ְŷ�ó DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtOrderCustomer, True)
            db = Filtered_DB(db, Me.txtProductionCustomer.Value, 1, False)
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

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
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


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub

