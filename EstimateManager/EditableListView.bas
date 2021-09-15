Attribute VB_Name = "EditableListView"
Option Explicit

Dim mouseX As Integer
Dim headerIndex As Integer
Dim beforeSelectedItem As ListItem
Dim currentEditText, currentCboText As Variant

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '����ID�� �ش��ϴ� ���� ������ �о��
    db = Get_DB(shtOrder)
    If Not isEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.value, 28, True)
    End If
    If Not isEmpty(db) Then
        db = Filtered_DB(db, "<>" & "����", 4)
    End If
    
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
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
        .SmallIcons = Me.ImageList1
        .Sorted = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_����", 0
        .ColumnHeaders.Add , , "������ȣ", 0
        .ColumnHeaders.Add , , "�з�", 34
        .ColumnHeaders.Add , , "�ŷ�ó", 50
        .ColumnHeaders.Add , , "ǰ��", 115
        .ColumnHeaders.Add , , "����", 60
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
        .ColumnHeaders.Add , , "����", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "����", 30
        
        '.ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not isEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))   'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_����
                li.ListSubItems.Add , , db(i, 5)        '������ȣ
                li.ListSubItems.Add , , db(i, 4)        '�з�
                li.ListSubItems.Add , , db(i, 6)        '�ŷ�ó
                li.ListSubItems.Add , , db(i, 7)        'ǰ��
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
                li.ListSubItems.Add , , db(i, 22)       '����
                li.ListSubItems.Add , , "����"       '����
                li.Selected = False
                
                If IsNumeric(db(i, 13)) Then
                    '��� �հ� ����
                    totalCost = totalCost + CLng(db(i, 13))
                End If
            Next
        End If
        
        If totalCost <> 0 Then
            Me.txtExecutionCost.value = Format(totalCost, "#,##0")
        End If
    End With
End Sub

Private Sub lswOrderList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * x
End Sub

Private Sub lswOrderList_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.value = ""
    Me.cboOrderCategory.value = ""
End Sub

Private Sub lswOrderList_DblClick()
    Dim i As Integer
    Dim pos As Integer
    
    bFocusOrderList = True
    bFocusPaymentList = False
    
    With Me.lswOrderList
        headerIndex = 0
        For i = 1 To .ColumnHeaders.count
            pos = .ColumnHeaders(i).Left
            If mouseX < pos Then
                headerIndex = i - 1
                Exit For
            End If
        Next
        
        If headerIndex = 0 Then
            If Not .selectedItem Is Nothing Then
                clickOrderId = .selectedItem.Text
                
                If isFormLoaded("frmOrderUpdate") = True Then
                    Unload frmOrderUpdate
                End If
                frmOrderUpdate.Show (False)
            End If
        ElseIf headerIndex = 12 Then
            '�ݾ��� ������ �� ����
        Else
            ' ���� ������ ���� �����س���
            If Not beforeSelectedItem Is Nothing Then
                Set beforeSelectedItem = Nothing
            End If
            Set beforeSelectedItem = .selectedItem
            
            SelectOrderListColumn
        End If
    End With

End Sub

Private Sub lswOrderList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswOrderList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long
    
    With Me.lswOrderList
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
            '���氪�� DB�� ȭ�鿡 �ݿ�
            OrderListUpdate headerIndex
            
            '����Ű - ���� �ٲ���. ����ĭ���� �̵����� ����
            If KeyCode = vbKeyReturn Then
                Me.txtEdit.Visible = False
                Me.frmEdit.Visible = False
                .SetFocus
            ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyRight Then
                '��Ű, ������ ȭ��ǥŰ
                If headerIndex = 18 Then
                    Me.txtEdit.Visible = False
                    Me.frmEdit.Visible = False
                    .SetFocus
                ElseIf headerIndex = 11 Then
                    headerIndex = headerIndex + 2
                    SelectOrderListColumn
                    KeyCode = 0
                Else
                    headerIndex = headerIndex + 1
                    SelectOrderListColumn
                    KeyCode = 0
                End If
            ElseIf KeyCode = vbKeyUp Then
                '����ȭ��ǥŰ
                '����Ʈ �� ó���� �ƴϸ� ��ĭ���� �̵�
                With Me.lswOrderList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = 1 Then
                                Me.txtEdit.Visible = False
                                Me.frmEdit.Visible = False
                                .SetFocus
                            Else
                                .ListItems(i).Selected = False
                                .ListItems(i - 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectOrderListColumn
                                KeyCode = 0
                                Exit For
                            End If
                        End If
                    Next
                End With
            ElseIf KeyCode = vbKeyDown Then
                '�Ʒ�ȭ��ǥŰ
                With Me.lswOrderList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = .ListItems.count Then
                                '�� �������̸� ������
                                Me.txtEdit.Visible = False
                                Me.frmEdit.Visible = False
                                .SetFocus
                                Exit For
                            Else
                                '����Ʈ �� �������� �ƴϸ� ��ĭ �Ʒ��� �̵�
                                .ListItems(i).Selected = False
                                .ListItems(i + 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectOrderListColumn
                                Exit For
                            End If
                        End If
                    Next
                End With
                KeyCode = 0
            ElseIf KeyCode = vbKeyLeft Then
                '����ȭ��ǥŰ
                '�� ó���� �ƴϸ� ��ĭ �������� �̵�
                If headerIndex <= 4 Then
                    Me.txtEdit.Visible = False
                    Me.frmEdit.Visible = False
                    .SetFocus
                Else
                    If headerIndex = 13 Then
                        headerIndex = headerIndex - 2   '�ݾ� �ʵ� �ǳʶٱ� ���ؼ� -2 ����
                    Else
                        headerIndex = headerIndex - 1
                    End If
                    SelectOrderListColumn
                    KeyCode = 0
                End If
            End If
        
        ElseIf KeyCode = vbKeyEscape Then
            'ESCŰ
            Me.txtEdit.Visible = False
            Me.frmEdit.Visible = False
        End If
    End With
End Sub

Private Sub cboOrderCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswOrderList
        If KeyCode = vbKeyReturn Then
            If headerIndex = 0 Then headerIndex = 4
            OrderListUpdate headerIndex
            Me.cboOrderCategory.Visible = False
            Me.frmEdit.Visible = False
            .SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If headerIndex = 0 Then headerIndex = 4
            OrderListUpdate headerIndex
            headerIndex = headerIndex + 1
            SelectOrderListColumn
            KeyCode = 0
        End If
    End With
End Sub

Private Sub txtEdit_AfterUpdate()
    '��Ű�� ����Ű�� �ƴ� ���콺�� Ŭ���ؼ� ����� ���: currentEditText�� �����
    If headerIndex > 4 And headerIndex < Me.lswOrderList.ColumnHeaders.count Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.txtEdit.value <> currentEditText Then
                OrderListUpdate headerIndex
                headerIndex = 0
                currentEditText = ""
            End If
        End If
    End If
    
End Sub

Private Sub cboOrderCategory_AfterUpdate()
    If headerIndex = 4 Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.cboOrderCategory.value <> currentCboText Then
                OrderListUpdate headerIndex
                headerIndex = 0
                currentCboText = ""
            End If
        End If
    End If
End Sub

Sub SelectOrderListColumn()
    Dim ItemSel    As ListItem
    
    If Not lswOrderList.selectedItem Is Nothing Then
        If headerIndex = lswOrderList.ColumnHeaders.count Then
            frmEdit.Visible = False
            txtEdit.Visible = False
            cboOrderCategory.Visible = False
        End If
        
        Set ItemSel = lswOrderList.selectedItem
        ItemSel.EnsureVisible
            
        If headerIndex > 4 And headerIndex < lswOrderList.ColumnHeaders.count Then
            With frmEdit
                .Visible = True
                .top = ItemSel.top + lswOrderList.top
                .Left = lswOrderList.ColumnHeaders(headerIndex).Left + lswOrderList.Left
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.txtEdit
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Width = lswOrderList.ColumnHeaders(headerIndex).Width
                .Height = lswOrderList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentEditText = .Text
            End With
            Me.cboOrderCategory.Visible = False
        ElseIf headerIndex = 4 Then
            With frmEdit
                .Visible = True
                .top = ItemSel.top + lswOrderList.top
                .Left = lswOrderList.ColumnHeaders(headerIndex).Left + lswOrderList.Left
                .Width = cboOrderCategory.Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With cboOrderCategory
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Height = lswOrderList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentCboText = .Text
            End With
            Me.txtEdit.Visible = False
        End If
    End If

End Sub

Sub OrderListUpdate(headerIndex)
    Dim orderPrice As Long
    
    With Me.lswOrderList
        If .selectedItem Is Nothing Then
            Exit Sub
        End If
        
        If headerIndex = 4 Then
            If Me.cboOrderCategory.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.cboOrderCategory.value
                'DB ���̺� ����
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.cboOrderCategory.value
            End If
        Else
            If Me.txtEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '�Է°� ���� ����
                ConvertOrderListFormat Me.txtEdit, headerIndex
                '����Ʈ�� �� ����
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtEdit.value
                'DB ���̺� ����
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.txtEdit.value
                
                '����,�ܰ� ������ ��쿡�� �ݾ׵� �����ؾ� ��
                If headerIndex = 9 Or headerIndex = 11 Then
                    orderPrice = CalculateOrderListPrice(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(orderPrice, "#,##0")
                    UpdateOrderListValue .selectedItem.Text, 12, orderPrice
                End If
                '���డ �Ѿ� ���
                Me.txtExecutionCost = Format(CalculateOrderListTotalCost, "#,##0")
                CalculateAcceptedMargin
            End If
        End If
    End With
End Sub

Sub UpdateShtOrderHeaderIndex(orderId, headerIndex, value)
    Dim findRow, fieldNo As Long
    
    findRow = isExistInSheet(shtOrderAdmin.Range("B6"), orderId)
    If findRow <> 0 Then
        Select Case headerIndex
            Case 4  '�з�
                fieldNo = 7
            Case 5  '�ŷ�ó
                fieldNo = 8
            Case 6  'ǰ��
                fieldNo = 9
            Case 7  '����
                fieldNo = 10
            Case 8  '�԰�
                fieldNo = 11
            Case 9  '����
                fieldNo = 12
            Case 10  '����
                fieldNo = 13
            Case 11  '�ܰ�
                fieldNo = 14
            Case 12  '�ݾ�
                fieldNo = 15
            Case 13  '����
                fieldNo = 18
            Case 14  '����
                fieldNo = 19
            Case 15  '�԰�
                fieldNo = 20
            Case 16  '����
                fieldNo = 22
            Case 17  '��꼭
                fieldNo = 23
            Case 18  '����
                fieldNo = 24
        End Select
        
        shtOrderAdmin.Cells(findRow, fieldNo).value = value
    End If
End Sub

Private Sub btnOrderListInsert_Click()
    Dim lastId As Long
    Dim li As ListItem
    
    '���ָ���Ʈ�信 ���� �߰�
    Insert_Record shtOrder, _
                , , "����", currentManagementId, , , , , , , , , , _
                , , , , , _
                , , , , _
                , , _
                Date, , currentEstimateId, , False
    lastId = Get_LastID(shtOrder)
    
    With Me.lswOrderList
        Set li = .ListItems.Add(, , lastId)   'ID
        li.ListSubItems.Add , , currentEstimateId       'ID_����
        li.ListSubItems.Add , , currentManagementId        '������ȣ
        li.ListSubItems.Add , , "����"        '�з�
        li.ListSubItems.Add , , ""        '�ŷ�ó
        li.ListSubItems.Add , , ""        'ǰ��
        li.ListSubItems.Add , , ""        '����
        li.ListSubItems.Add , , ""        '�԰�
        li.ListSubItems.Add , , ""        '����
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , ""          '�ܰ�
        li.ListSubItems.Add , , ""      '�ݾ�
        li.ListSubItems.Add , , ""       '������
        li.ListSubItems.Add , , ""       '������
        li.ListSubItems.Add , , ""       '�԰���
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , ""       '��꼭
        li.ListSubItems.Add , , ""       '����
        li.ListSubItems.Add , , "����"       '����
        
        .selectedItem.Selected = False
        li.Selected = True
        li.EnsureVisible
        
        headerIndex = 4
        SelectOrderListColumn
    End With
End Sub

Sub UpdateOrderListValue(id, headerIndex, value)
    Dim fieldName As String

    Select Case headerIndex
        Case 4  '�з�
            fieldName = "�з�2"
        Case 5  '�ŷ�ó
            fieldName = "�ŷ�ó"
        Case 6  'ǰ��
            fieldName = "ǰ��"
        Case 7  '����
            fieldName = "����"
        Case 8  '�԰�
            fieldName = "�԰�"
        Case 9  '����
            fieldName = "����"
        Case 10  '����
            fieldName = "����"
        Case 11  '�ܰ�
            fieldName = "�ܰ�"
        Case 12  '�ݾ�
            fieldName = "�ݾ�"
        Case 13  '����
            fieldName = "����"
        Case 14  '����
            fieldName = "����"
        Case 15  '�԰�
            fieldName = "�԰�"
        Case 16  '����
            fieldName = "����"
        Case 17  '��꼭
            fieldName = "��꼭"
        Case 18  '����
            fieldName = "����"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtOrder, id, fieldName, value
        Update_Record_Column shtOrder, id, "��������", Date
    End If

End Sub


Private Sub Frame4_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub


Private Sub UserForm_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub


