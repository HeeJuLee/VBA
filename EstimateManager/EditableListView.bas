Attribute VB_Name = "EditableListView"
Option Explicit

Dim mouseX As Integer
Dim headerIndex As Integer
Dim beforeSelectedItem As ListItem
Dim currentEditText, currentCboText As Variant


Private Sub lswOrderList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * X
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


