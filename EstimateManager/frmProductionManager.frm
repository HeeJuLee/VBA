VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductionManager 
   Caption         =   "��������׸� ����"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   OleObjectBlob   =   "frmProductionManager.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmProductionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim estimate As Variant
    
    If currentEstimateId = "" Then
        MsgBox "currentEstimateId ����: ������ ������ �����ϴ�.", vbInformation, "�۾� Ȯ��"
        End
    End If
    
    '�ؽ�Ʈ�ڽ� �� ��Ʈ�� ���� ����
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            If contr.Name Like "lbl2*" Then
                'contr.BackColor = RGB(48, 84, 150)
                'contr.ForeColor = RGB(255, 255, 255)
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
        MsgBox "currentEstimateId�� �ش��ϴ� ���� �����Ͱ� �����ϴ�.", vbInformation, "�۾� Ȯ��"
        End
    End If

    Me.txtEstimateName.value = estimate(6)
    Me.txtManagementID.value = estimate(2)
    Me.txtCustomer.value = estimate(4)
    Me.txtManager.value = estimate(5)
    
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
    
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '����Ʈ�� �� ����
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwAutomatic
        .CheckBoxes = False
        .SmallIcons = Me.ImageList1
        
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
            
            Me.txtProductionTotalCost.value = Format(totalCost, "#,##0")
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
    
    If Me.txtProductionItem.value = "" Then MsgBox "ǰ���� �Է��ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "�ݾ��� �Է��ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub

    '��������׸� ����
    Insert_Record shtProduction, CLng(currentEstimateId), Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Me.cboCategory.value, Date
    
    '������డ�� ������ ���� ������Ʈ
    RefreshProductionTotalCost
    
    '����� ������ ����
    Me.txtProductionID.value = Get_LastID(shtProduction)
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub


Sub UpdateProduction()
    Dim cost As Variant

    If Me.txtProductionID.value = "" Then MsgBox "������ �׸��� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    If Me.txtProductionItem.value = "" Then MsgBox "ǰ���� �Է��ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "�ݾ��� �Է��ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    '���� ��������׸� ������Ʈ
    Update_Record shtProduction, Me.txtProductionID.value, currentEstimateId, Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Me.cboCategory.value, Date
    
    '������డ�� ������ ���� ������Ʈ
    RefreshProductionTotalCost
    
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub

Sub RefreshProductionTotalCost()
    '��������׸� �հ� ���
    Me.txtProductionTotalCost.value = Format(GetProductionTotalCost, "#,##0")
    
    '������డ�� �������̺� ����
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "���డ(����)", CLng(Me.txtProductionTotalCost.value)
    
    '������డ�� frmEstimateUpdate �� ���� ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        'frmEstimateUpdate.txtProductionTotalCost = Me.txtProductionTotalCost.value
        frmEstimateUpdate.UpdateProductionTotalCost Me.txtProductionTotalCost.value
'        frmEstimateUpdate.CalculateEstimateUpdateCost
'        frmEstimateUpdate.UpdateShtEstimateField currentEstimateId, "������డ", Me.txtProductionTotalCost.value
    End If
    
    '��������׸� ����Ʈ�ڽ� ���ΰ�ħ
    InitializeLswProductionList
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
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.", vbInformation, "�۾� Ȯ��": Exit Sub
    
    YN = MsgBox("������ " & count & "�� �׸��� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            '��������׸� ���̺��� ����
            Delete_Record shtProduction, li.SubItems(1)
        End If
    Next
    
    '������డ�� ������ ���� ������Ʈ
    RefreshProductionTotalCost
        
    Me.txtProductionID.value = ""
    ClearProductionInput
    
End Sub

Sub ProductionToOrder()
    Dim li As ListItem
    Dim count As Long
    Dim managementId, category, customer, Item, material, size, amount, unit, unitPrice, cost, memo As Variant
    Dim YN As VbMsgBoxResult
    Dim estimate As Variant
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    '���� Ȯ���� �ƴ� ������ ���� ���ָ� �� �� ����
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If estimate(38) = "" Then
        MsgBox "���� Ȯ���� �����ؾ� ������ �� �ֽ��ϴ�.", vbInformation, "�۾� Ȯ��"
        Exit Sub
    End If
    
    YN = MsgBox("������ " & count & "�� �׸��� �����ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If YN = vbNo Then Exit Sub
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            Item = li.Text
            managementId = li.SubItems(3)
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
                , , category, managementId, customer, Item, material, size, amount, unit, unitPrice, cost, , _
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
    
    MsgBox "�� " & count & "�� �׸��� �����Ͽ����ϴ�.", vbInformation, "�۾� Ȯ��"
    
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
                    .SetFocus
                    .selectedItem = .ListItems(i)
                    .selectedItem.EnsureVisible
                Else
                    .ListItems(i).Selected = False
                End If
            Next
        End If
    End With
End Sub

Sub ClearProductionInput()
    Me.txtProductionID.value = ""
    Me.txtProductionCustomer.value = ""
    Me.txtProductionItem.value = ""
    Me.txtProductionMaterial.value = ""
    Me.txtProductionSize.value = ""
    Me.txtProductionAmount.value = ""
    Me.cboProductionUnit.value = ""
    Me.txtProductionUnitPrice.value = ""
    Me.txtProductionCost.value = ""
    Me.txtProductionMemo.value = ""
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

Private Sub btnProductionCopy_Click()
    If isFormLoaded("frmProductionCopy") Then
        Unload frmProductionCopy
    End If
    frmProductionCopy.Show (False)
End Sub

Private Sub btnProductionClose_Click()
    Unload Me
End Sub

Private Sub lswProductionList_Click()
    With Me.lswProductionList
        If Not .selectedItem Is Nothing Then
            Me.txtProductionItem.value = .selectedItem.Text
            Me.txtProductionID.value = .selectedItem.ListSubItems(1)
            Me.cboCategory.value = .selectedItem.ListSubItems(4)
            Me.txtProductionCustomer.value = .selectedItem.ListSubItems(5)
            Me.txtProductionMaterial.value = .selectedItem.ListSubItems(6)
            Me.txtProductionSize.value = .selectedItem.ListSubItems(7)
            Me.txtProductionAmount.value = .selectedItem.ListSubItems(8)
            Me.cboProductionUnit.value = .selectedItem.ListSubItems(9)
            Me.txtProductionUnitPrice.value = .selectedItem.ListSubItems(10)
            Me.txtProductionCost.value = .selectedItem.ListSubItems(11)
            Me.txtProductionMemo.value = .selectedItem.ListSubItems(12)
        End If
    End With
End Sub

Private Sub lswProductionList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswProductionList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub


Private Sub btnProductionClear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnProductionClose.SetFocus
    End If
End Sub

Private Sub cboCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtProductionCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '����Ű - ���� �Է�ĭ���� �̵�
        Me.lswOrderCustomerAutoComplete.Visible = False
        Me.txtProductionItem.SetFocus
    ElseIf KeyCode = 9 Then
        '��Ű �ڵ��ϼ��� �ϳ��̸� �������� �̵�
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count = 1 Then
                Me.lswOrderCustomerAutoComplete.Visible = False
                Me.txtProductionItem.SetFocus
                KeyCode = 0
            Else
                If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
            End If
        End With
    ElseIf KeyCode = 40 Then
        '�Ʒ�ȭ��Ű - �ڵ��ϼ� ����� �ִ� ��쿡�� ��Ŀ���� �ڵ��ϼ� ����Ʈ�� �̵�
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
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
        If Me.txtProductionCustomer.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '���ְŷ�ó DB�� �о�ͼ� ����Ʈ�信 ���
            .ListItems.Clear
            db = Get_DB(shtOrderCustomer, True)
            db = Filtered_DB(db, Me.txtProductionCustomer.value, 1, False)
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
        If Not .selectedItem Is Nothing Then
            Me.txtProductionCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtProductionItem.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '�ŷ�ó�� ���� �־��ְ� ��Ŀ���� ǰ������ �̵�
    If KeyCode = 13 Then
        With Me.lswOrderCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtProductionCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtProductionItem.SetFocus
            End If
        End With
    End If
End Sub


Private Sub txtProductionItem_Enter()
    If Me.lswOrderCustomerAutoComplete.Visible = True Then
        With Me.lswOrderCustomerAutoComplete
            Me.txtProductionCustomer.value = .selectedItem.Text
            .Visible = False
        End With
    End If
End Sub


Private Sub txtProductionCustomer_AfterUpdate()
    Me.txtProductionCustomer.value = Trim(Me.txtProductionCustomer.value)
End Sub


Private Sub cboCategory_AfterUpdate()
    Me.cboCategory.value = Trim(Me.cboCategory.value)
End Sub


Private Sub txtProductionAmount_AfterUpdate()
    If Me.txtProductionAmount.value = "" Then
        Me.txtProductionCost.value = Me.txtProductionUnitPrice.value
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionAmount.value) Then
        MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        Me.txtProductionAmount.value = ""
        Me.txtProductionCost.value = Me.txtProductionUnitPrice.value
        Exit Sub
    End If
        
    Me.txtProductionAmount.value = Format(Me.txtProductionAmount.value, "#,##0")
        
    '�ݾ� = ���� * �ܰ�
    If IsNumeric(Me.txtProductionUnitPrice.value) Then
        Me.txtProductionCost.value = Format(CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value), "#,##0")
    End If
End Sub


Private Sub txtProductionItem_AfterUpdate()
    Me.txtProductionItem.value = Trim(Me.txtProductionItem.value)
End Sub

Private Sub txtProductionMaterial_AfterUpdate()
    Me.txtProductionMaterial.value = Trim(Me.txtProductionMaterial.value)
End Sub


Private Sub txtProductionMemo_AfterUpdate()
    Me.txtProductionMemo.value = Trim(Me.txtProductionMemo.value)
End Sub


Private Sub txtProductionSize_AfterUpdate()
    Me.txtProductionSize.value = Trim(Me.txtProductionSize.value)
End Sub


Private Sub txtProductionUnitPrice_AfterUpdate()
    If Me.txtProductionUnitPrice.value = "" Then
        Me.txtProductionCost.value = ""
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionUnitPrice.value) Then
        MsgBox "���ڸ� �Է��ϼ���.", vbInformation, "�۾� Ȯ��"
        Me.txtProductionUnitPrice.value = ""
        Me.txtProductionCost.value = ""
        Exit Sub
    End If
    
    Me.txtProductionUnitPrice.value = Format(Me.txtProductionUnitPrice.value, "#,##0")
    
    If IsNumeric(Me.txtProductionUnitPrice.value) Then
        If Me.txtProductionAmount.value = "" Then
            Me.txtProductionCost.value = Format(Me.txtProductionCost.value, "#,##0")
        Else
            If IsNumeric(Me.txtProductionAmount.value) Then
                '�ݾ� = ���� * �ܰ�
                Me.txtProductionCost.value = Format(CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value), "#,##0")
            End If
        End If
    End If
    
End Sub

Private Sub UserForm_Layout()
    productionFormX = Me.Left
    productionFormY = Me.top
End Sub

