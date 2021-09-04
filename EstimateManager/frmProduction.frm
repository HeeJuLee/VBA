VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProduction 
   Caption         =   "��������׸� ����"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   OleObjectBlob   =   "frmProduction.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnProductionClear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnProductionClose.SetFocus
    End If
End Sub

Private Sub btnProductionClose_Click()
    Unload Me
End Sub

Private Sub cboCategory_AfterUpdate()
    Me.cboCategory.Value = Trim(Me.cboCategory.Value)
End Sub


Private Sub lswProductionList_Click()
    With Me.lswProductionList
        If Not .SelectedItem Is Nothing Then
            Me.txtProductionItem.Value = .SelectedItem.Text
            Me.txtProductionID.Value = .SelectedItem.ListSubItems(1)
            Me.cboCategory.Value = .SelectedItem.ListSubItems(4)
            Me.txtProductionCustomer.Value = .SelectedItem.ListSubItems(5)
            Me.txtProductionMaterial.Value = .SelectedItem.ListSubItems(6)
            Me.txtProductionSize.Value = .SelectedItem.ListSubItems(7)
            Me.txtProductionAmount.Value = .SelectedItem.ListSubItems(8)
            Me.cboProductionUnit.Value = .SelectedItem.ListSubItems(9)
            Me.txtProductionUnitPrice.Value = .SelectedItem.ListSubItems(10)
            Me.txtProductionCost.Value = .SelectedItem.ListSubItems(11)
            Me.txtProductionMemo.Value = .SelectedItem.ListSubItems(12)
        End If
    End With
End Sub




Private Sub txtProductionCustomer_AfterUpdate()
    Me.txtProductionCustomer.Value = Trim(Me.txtProductionCustomer.Value)
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
    ElseIf KeyCode = 27 Then
        'ESCŰ �ݱ�
        Unload Me
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

Private Sub txtProductionAmount_AfterUpdate()
    If Me.txtProductionAmount.Value = "" Then
        Me.txtProductionCost.Value = Me.txtProductionUnitPrice.Value
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionAmount.Value) Then
        MsgBox "���ڸ� �Է��ϼ���."
        Me.txtProductionAmount.Value = ""
        Me.txtProductionCost.Value = Me.txtProductionUnitPrice.Value
        Exit Sub
    End If
        
    Me.txtProductionAmount.Value = Format(Me.txtProductionAmount.Value, "#,##0")
        
    '�ݾ� = ���� * �ܰ�
    If IsNumeric(Me.txtProductionUnitPrice.Value) Then
        Me.txtProductionCost.Value = Format(CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value), "#,##0")
    End If
End Sub

Private Sub txtProductionItem_AfterUpdate()
    Me.txtProductionItem.Value = Trim(Me.txtProductionItem.Value)
End Sub


Private Sub txtProductionMaterial_AfterUpdate()
    Me.txtProductionMaterial.Value = Trim(Me.txtProductionMaterial.Value)
End Sub


Private Sub txtProductionMemo_AfterUpdate()
    Me.txtProductionMemo.Value = Trim(Me.txtProductionMemo.Value)
End Sub


Private Sub txtProductionSize_AfterUpdate()
    Me.txtProductionSize.Value = Trim(Me.txtProductionSize.Value)
End Sub


Private Sub txtProductionUnitPrice_AfterUpdate()
    If Me.txtProductionUnitPrice.Value = "" Then
        Me.txtProductionCost.Value = ""
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionUnitPrice.Value) Then
        MsgBox "���ڸ� �Է��ϼ���."
        Me.txtProductionUnitPrice.Value = ""
        Me.txtProductionCost.Value = ""
        Exit Sub
    End If
    
    Me.txtProductionUnitPrice.Value = Format(Me.txtProductionUnitPrice.Value, "#,##0")
    
    If IsNumeric(Me.txtProductionUnitPrice.Value) Then
        If Me.txtProductionAmount.Value = "" Then
            Me.txtProductionCost.Value = Format(Me.txtProductionCost.Value, "#,##0")
        Else
            If IsNumeric(Me.txtProductionAmount.Value) Then
                '�ݾ� = ���� * �ܰ�
                Me.txtProductionCost.Value = Format(CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value), "#,##0")
            End If
        End If
    End If
    
End Sub


Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim estimate As Variant
    
    If currentEstimateId = "" Then
        MsgBox "currentEstimateId ����: ������ ������ �����ϴ�."
        End
    End If
    
    '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            If contr.Name Like "lbl2*" Then
                contr.BackColor = RGB(48, 84, 150)
                contr.ForeColor = RGB(255, 255, 255)
            ElseIf contr.Name Like "lbl3*" Then
                contr.BackColor = RGB(221, 235, 247)
            Else
                contr.BackColor = RGB(242, 242, 242)
            End If
        End If
    Next
    
    '�� ��ġ ����
    If productionFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = productionFormX
        Me.top = productionFormY
    End If
    
    'currentEstimateId�� ���������� �о���� (Ȯ�ο�)
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If IsEmpty(estimate) Then
        MsgBox "currentEstimateId�� �ش��ϴ� ���� �����Ͱ� �����ϴ�."
        End
    End If

    Me.txtEstimateName.Value = estimate(6)
    Me.txtManagementID.Value = estimate(2)
    Me.txtCustomer.Value = estimate(4)
    Me.txtManager.Value = estimate(5)
    
    InitializeCboCategory           '�з�
    InitializeLswProductionList    '��������׸� ���
    InitializeCboProductonUnit  '��������׸� ����
    InitializeLswOrderCustomerAutoComplete   '���ְŷ�ó �ڵ��ϼ�
    
    ClearProductionInput
End Sub

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '����ID�� �ش��ϴ� �������׸��� �о��
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwManual
        .CheckBoxes = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ǰ��", 130
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 70
        .ColumnHeaders.Add , , "����", 60
        .ColumnHeaders.Add , , "�԰�", 80
        .ColumnHeaders.Add , , "����", 44, lvwColumnRight
        .ColumnHeaders.Add , , "����", 44, lvwColumnCenter
        .ColumnHeaders.Add , , "�ܰ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�ݾ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�޸�", 92
        .ColumnHeaders.Add , , "�������", 0
        
        .ColumnHeaders(1).Position = 6
    
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
                li.ListSubItems.Add , , db(i, 13)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , db(i, 8)
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                
                li.Selected = False
            Next
            
            Me.txtProductionTotalCost.Value = Format(totalCost, "#,##0")
        End If
    End With
End Sub

Sub InitializeLswOrderCustomerAutoComplete()
    With Me.lswOrderCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeCboCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboCategory, db
End Sub

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
End Sub


Sub InsertProduction()
    
    If Me.txtProductionItem.Value = "" Then MsgBox "ǰ���� �Է��ϼ���.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "�ݾ��� �Է��ϼ���.": Exit Sub

    '��������׸� ����
    Insert_Record shtProduction, CLng(currentEstimateId), Me.txtManagementID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Me.cboCategory.Value, Date
    
    '��������׸� �հ� ���
    Me.txtProductionTotalCost.Value = Format(GetProductionTotalCost, "#,##0")
    
    '������డ, ��������, ������, �������� �ݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "���డ(����)", CLng(Me.txtProductionTotalCost.Value)
    
    '������డ�� frmEstimateUpdate �� ���� ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtProductionTotalCost = Me.txtProductionTotalCost.Value
        frmEstimateUpdate.CalculateEstimateUpdateCost
    End If
    
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
    
    '���� ��������׸� ������Ʈ
    Update_Record shtProduction, Me.txtProductionID.Value, currentEstimateId, Me.txtManagementID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Me.cboCategory.Value, Date
    
    '������డ �հ� ���
    Me.txtProductionTotalCost.Value = Format(GetProductionTotalCost, "#,##0")
    
    '������డ�� �������̺� ����
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "���డ(����)", CLng(Me.txtProductionTotalCost.Value)
    
    '������డ�� frmEstimateUpdate �� ���� ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtProductionTotalCost = Me.txtProductionTotalCost.Value
        frmEstimateUpdate.CalculateEstimateUpdateCost
    End If
    
    InitializeLswProductionList
    SelectItemLswProduction Me.txtProductionID.Value
    
End Sub


Sub DeleteProduction()
    Dim db As Variant
    Dim YN As VbMsgBoxResult
    Dim count As Long
    Dim li As ListItem

    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    YN = MsgBox("������ " & count & "�� �׸��� �����մϴ�.", vbYesNo)
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            '��������׸� ���̺��� ����
            Delete_Record shtProduction, li.SubItems(1)
        End If
    Next
    
    If count > 0 Then
        '������డ ���
        Me.txtProductionTotalCost.Value = Format(GetProductionTotalCost, "#,##0")
        
         '������డ, ��������, ������, �������� �ݾ��� �������̺� ����
        Update_Record_Column shtEstimate, CLng(currentEstimateId), "���డ(����)", CLng(Me.txtProductionTotalCost.Value)
        
        '������డ�� frmEstimateUpdate �� ���� ������Ʈ
        If isFormLoaded("frmEstimateUpdate") Then
            frmEstimateUpdate.txtProductionTotalCost = Me.txtProductionTotalCost.Value
            frmEstimateUpdate.CalculateEstimateUpdateCost
        End If
    End If
        
    Me.txtProductionID.Value = ""
    InitializeLswProductionList
    ClearProductionInput
    
End Sub

Sub ProductionToOrder()
    Dim li As ListItem
    Dim count As Long
    Dim managementID, category, customer, Item, material, size, amount, unit, unitPrice, cost, memo As Variant
    Dim YN As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    YN = MsgBox("������ " & count & "�� �׸��� �����մϴ�.", vbYesNo)
    If YN = vbNo Then Exit Sub
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            Item = li.Text
            managementID = li.SubItems(3)
            category = li.SubItems(4)
            customer = li.SubItems(5)
            material = li.SubItems(6)
            size = li.SubItems(7)
            amount = li.SubItems(8)
            unit = li.SubItems(9)
            unitPrice = li.SubItems(10)
            cost = li.SubItems(11)
            memo = li.SubItems(12)
            
            '������ ��������׸��� ���� ���̺� ���
            Insert_Record shtOrder, _
                , , category, managementID, customer, Item, material, size, amount, unit, unitPrice, cost, , _
                , , , , , _
                , , , , _
                , , _
                Date, , currentEstimateId, memo, False
                
            count = count + 1
        End If
    Next
    
    'frmEstimateUpdate ���� ���ָ���� ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.InitializeLswOrderList
        frmEstimateUpdate.CalculateEstimateUpdateCost
    End If
    
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
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
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

Sub SelectItemLswProduction(selectedID As Variant)
    Dim i As Long
    
    With Me.lswProductionList
        If Not IsMissing(selectedID) Then
            For i = 1 To .ListItems.count
                If selectedID = .ListItems(i).SubItems(1) Then
                    .SelectedItem = .ListItems(i)
                    .SetFocus
                Else
                    .ListItems(i).Selected = False
                End If
            Next
        End If
    End With
End Sub

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

Private Sub UserForm_Layout()
    productionFormX = Me.Left
    productionFormY = Me.top
End Sub

