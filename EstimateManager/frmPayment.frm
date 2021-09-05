VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPayment 
   Caption         =   "���� �̷� ����"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "frmPayment.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private acceptedPrice As String

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
    If paymentFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = paymentFormX
        Me.top = paymentFormY
    End If
    
    'currentEstimateId�� ���������� �о���� (Ȯ�ο�)
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If IsEmpty(estimate) Then
        MsgBox "currentEstimateId�� �ش��ϴ� ���� �����Ͱ� �����ϴ�."
        End
    End If

    Me.txtEstimateName.Value = estimate(6)
    Me.txtManagementID.Value = estimate(2)
    acceptedPrice = estimate(21)
    
    InitializeLswPaymentList    '���� �̷�
    
    ClearPaymentInput
End Sub

Sub InitializeLswPaymentList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '����ID�� �ش��ϴ� �����̷��� �о��
    db = Get_DB(shtPayment)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
     '����Ʈ�� �� ����
    With Me.lswPaymentList
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
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "������", 70, lvwColumnCenter
        .ColumnHeaders.Add , , "�����ݾ�", 70, lvwColumnRight
        .ColumnHeaders.Add , , "�޸�", 140
        .ColumnHeaders.Add , , "�������", 0
        
        '.ColumnHeaders(1).Position = 1
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 5)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 5))
                End If
                
                Set li = .ListItems.Add(, , db(i, 1))
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , Format(db(i, 5), "#,##0")
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                
                li.Selected = False
            Next
            
            Me.txtPaid.Value = Format(totalCost, "#,##0")
            If IsNumeric(acceptedPrice) Then
                Me.txtRemaining.Value = Format(acceptedPrice - totalCost, "#,##0")
            End If
        End If
    End With
End Sub

Sub InsertPayment()
    
    If Me.txtPayDate.Value = "" Then MsgBox "�������� �Է��ϼ���.": Exit Sub
    If Me.txtPayAmount.Value = "" Then MsgBox "�����ݾ��� �Է��ϼ���.": Exit Sub

    '���Ұ����̷¿� ����
    Insert_Record shtPayment, CLng(currentEstimateId), Me.txtManagementID.Value, Me.txtPayDate.Value, Me.txtPayAmount.Value, Me.txtPayMemo.Value, Date
    
    '�հ� ���
    Me.txtPaid.Value = Format(GetPaymentTotalCost, "#,##0")
    If IsNumeric(acceptedPrice) Then
        Me.txtRemaining.Value = Format(acceptedPrice - Me.txtPaid.Value, "#,##0")
    End If
    
    '�Աݾ�/���Աݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "�Աݾ�", Me.txtPaid.Value
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "���Աݾ�", Me.txtRemaining.Value
    
    '�Աݾ�/���Աݾ��� frmEstimateUpdate ������ ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtPaymentDate.Value = Me.txtPayDate.Value
        frmEstimateUpdate.txtPaid.Value = Me.txtPaid.Value
        frmEstimateUpdate.txtRemaining.Value = Me.txtRemaining.Value
    End If
    
    InitializeLswPaymentList
    
    '����� ������ ����
    Me.txtPayID.Value = Get_LastID(shtPayment)
    SelectItemLswPayment Me.txtPayID.Value
    
End Sub


Sub UpdatePayment()
    Dim cost As Variant

    If Me.txtPayID.Value = "" Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    If Me.txtPayDate.Value = "" Then MsgBox "�������� �Է��ϼ���.": Exit Sub
    If Me.txtPayAmount.Value = "" Then MsgBox "�����ݾ��� �Է��ϼ���.": Exit Sub
    
    '���� ���Ұ����̷¿� ������Ʈ
    Update_Record shtPayment, Me.txtPayID.Value, currentEstimateId, Me.txtManagementID.Value, Me.txtPayDate.Value, Me.txtPayAmount.Value, Me.txtPayMemo.Value, Date
    
    '�հ� ���
    Me.txtPaid.Value = Format(GetPaymentTotalCost, "#,##0")
    If IsNumeric(acceptedPrice) Then
        Me.txtRemaining.Value = Format(acceptedPrice - Me.txtPaid.Value, "#,##0")
    End If
    
    '�Աݾ�/���Աݾ��� �������̺� ����
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "�Աݾ�", Me.txtPaid.Value
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "���Աݾ�", Me.txtRemaining.Value
    
    '�Աݾ�/���Աݾ��� frmEstimateUpdate ������ ������Ʈ
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtPaymentDate.Value = Me.txtPayDate.Value
        frmEstimateUpdate.txtPaid.Value = Me.txtPaid.Value
        frmEstimateUpdate.txtRemaining.Value = Me.txtRemaining.Value
    End If
    
    InitializeLswPaymentList
    SelectItemLswPayment Me.txtPayID.Value
    
End Sub


Sub DeletePayment()
    Dim db As Variant
    Dim YN As VbMsgBoxResult
    Dim count As Long
    Dim li As ListItem

    count = 0
    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "������ �׸��� �����ϼ���.": Exit Sub
    
    YN = MsgBox("������ " & count & "�� �׸��� �����մϴ�.", vbYesNo)
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            '�����̷� ���̺��� ����
            Delete_Record shtPayment, li.Text
        End If
    Next
    
    If count > 0 Then
        '�հ� ���
        Me.txtPaid.Value = Format(GetPaymentTotalCost, "#,##0")
        If IsNumeric(acceptedPrice) Then
            Me.txtRemaining.Value = Format(acceptedPrice - Me.txtPaid.Value, "#,##0")
        End If
        
        '�Աݾ�/���Աݾ��� �������̺� ����
        Update_Record_Column shtEstimate, CLng(currentEstimateId), "�Աݾ�", Me.txtPaid.Value
        Update_Record_Column shtEstimate, CLng(currentEstimateId), "���Աݾ�", Me.txtRemaining.Value
        
        '�Աݾ�/���Աݾ��� frmEstimateUpdate ������ ������Ʈ
        If isFormLoaded("frmEstimateUpdate") Then
            frmEstimateUpdate.txtPaymentDate.Value = Me.txtPayDate.Value
            frmEstimateUpdate.txtPaid.Value = Me.txtPaid.Value
            frmEstimateUpdate.txtRemaining.Value = Me.txtRemaining.Value
        End If
    End If
        
    Me.txtPayID.Value = ""
    InitializeLswPaymentList
    ClearPaymentInput
    
End Sub
Function GetPaymentTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '����ID�� �ش��ϴ� �����̷��� �о��
    db = Get_DB(shtPayment)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
    'DB�� ���� ���� ���
    totalCost = 0
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 5)) Then
                '��� �հ� ����
                totalCost = totalCost + CLng(db(i, 5))
            End If
        Next
    End If
        
    GetPaymentTotalCost = totalCost
End Function

Sub SelectItemLswPayment(selectedID As Variant)
    Dim i As Long
    
    With Me.lswPaymentList
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

Sub ClearPaymentInput()
    Me.txtPayID.Value = ""
    Me.txtPayDate.Value = ""
    Me.txtPayAmount.Value = ""
    Me.txtPayMemo.Value = ""
End Sub

Private Sub btnPaymentClear_Click()
    ClearPaymentInput
End Sub

Private Sub btnPaymentDelete_Click()
    DeletePayment
End Sub

Private Sub btnPaymentInsert_Click()
    InsertPayment
End Sub

Private Sub btnPaymentUpdate_Click()
    UpdatePayment
End Sub


Private Sub btnPaymentClose_Click()
    Unload Me
End Sub

Private Sub lswPaymentList_Click()
    With Me.lswPaymentList
        If Not .SelectedItem Is Nothing Then
            Me.txtPayID.Value = .SelectedItem.Text
            Me.txtPayDate.Value = .SelectedItem.ListSubItems(1)
            Me.txtPayAmount.Value = .SelectedItem.ListSubItems(2)
            Me.txtPayMemo.Value = .SelectedItem.ListSubItems(3)
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

Private Sub btnPaymentClear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnPaymentClose.SetFocus
    End If
End Sub

Private Sub imgPayDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPayDate
End Sub


Private Sub txtPayAmount_AfterUpdate()
    Me.txtPayAmount.Value = Trim(Me.txtPayAmount.Value)
    
    If Not IsNumeric(Me.txtPayAmount.Value) Then
        MsgBox "���ڸ� �Է��ϼ���.", vbExclamation
        Exit Sub
    End If
    
    Me.txtPayAmount.Value = Format(Me.txtPayAmount.Value, "#,##0")
End Sub

Private Sub txtPayDate_AfterUpdate()
    Me.txtPayDate.Value = Trim(Me.txtPayDate.Value)
End Sub

Private Sub txtPayDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtPayMemo_AfterUpdate()
    Me.txtPayMemo.Value = Trim(Me.txtPayMemo.Value)
End Sub


Private Sub UserForm_Layout()
    paymentFormX = Me.Left
    paymentFormY = Me.top
End Sub

