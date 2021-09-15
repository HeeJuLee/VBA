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
    
    '견적ID에 해당하는 발주 정보를 읽어옴
    db = Get_DB(shtOrder)
    If Not isEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.value, 28, True)
    End If
    If Not isEmpty(db) Then
        db = Filtered_DB(db, "<>" & "수주", 4)
    End If
    
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '리스트뷰 값 설정
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
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "분류", 34
        .ColumnHeaders.Add , , "거래처", 50
        .ColumnHeaders.Add , , "품목", 115
        .ColumnHeaders.Add , , "재질", 60
        .ColumnHeaders.Add , , "규격", 62
        .ColumnHeaders.Add , , "수량", 30, lvwColumnRight
        .ColumnHeaders.Add , , "단위", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "단가", 60, lvwColumnRight
        .ColumnHeaders.Add , , "금액", 60, lvwColumnRight
        .ColumnHeaders.Add , , "발주", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "납기", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "입고", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "명세서", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "계산서", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "결제", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "수정", 30
        
        '.ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not isEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))   'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_견적
                li.ListSubItems.Add , , db(i, 5)        '관리번호
                li.ListSubItems.Add , , db(i, 4)        '분류
                li.ListSubItems.Add , , db(i, 6)        '거래처
                li.ListSubItems.Add , , db(i, 7)        '품목
                li.ListSubItems.Add , , db(i, 8)        '재질
                li.ListSubItems.Add , , db(i, 9)        '규격
                li.ListSubItems.Add , , db(i, 10)        '수량
                li.ListSubItems.Add , , db(i, 11)       '단위
                li.ListSubItems.Add , , Format(db(i, 12), "#,##0")      '단가
                li.ListSubItems.Add , , Format(db(i, 13), "#,##0")      '금액
                li.ListSubItems.Add , , db(i, 16)       '발주일
                li.ListSubItems.Add , , db(i, 17)       '납기일
                li.ListSubItems.Add , , db(i, 18)       '입고일
                li.ListSubItems.Add , , db(i, 20)       '명세서
                li.ListSubItems.Add , , db(i, 21)       '계산서
                li.ListSubItems.Add , , db(i, 22)       '결제
                li.ListSubItems.Add , , "열기"       '수정
                li.Selected = False
                
                If IsNumeric(db(i, 13)) Then
                    '비용 합계 구함
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
            '금액은 변경할 수 없음
        Else
            ' 현재 선택한 열을 저장해놓음
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
            '변경값을 DB와 화면에 반영
            OrderListUpdate headerIndex
            
            '엔터키 - 값만 바꿔줌. 다음칸으로 이동하지 않음
            If KeyCode = vbKeyReturn Then
                Me.txtEdit.Visible = False
                Me.frmEdit.Visible = False
                .SetFocus
            ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyRight Then
                '탭키, 오른쪽 화살표키
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
                '위쪽화살표키
                '리스트 맨 처음이 아니면 한칸위로 이동
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
                '아래화살표키
                With Me.lswOrderList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = .ListItems.count Then
                                '맨 마지막이면 마무리
                                Me.txtEdit.Visible = False
                                Me.frmEdit.Visible = False
                                .SetFocus
                                Exit For
                            Else
                                '리스트 맨 마지막이 아니면 한칸 아래로 이동
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
                '왼쪽화살표키
                '맨 처음이 아니면 한칸 왼쪽으로 이동
                If headerIndex <= 4 Then
                    Me.txtEdit.Visible = False
                    Me.frmEdit.Visible = False
                    .SetFocus
                Else
                    If headerIndex = 13 Then
                        headerIndex = headerIndex - 2   '금액 필드 건너뛰기 위해서 -2 해줌
                    Else
                        headerIndex = headerIndex - 1
                    End If
                    SelectOrderListColumn
                    KeyCode = 0
                End If
            End If
        
        ElseIf KeyCode = vbKeyEscape Then
            'ESC키
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
    '탭키나 엔터키가 아닌 마우스를 클릭해서 벗어나는 경우: currentEditText를 사용함
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
                '리스트뷰 값 변경
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.cboOrderCategory.value
                'DB 테이블 변경
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.cboOrderCategory.value
            End If
        Else
            If Me.txtEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '입력값 포맷 변경
                ConvertOrderListFormat Me.txtEdit, headerIndex
                '리스트뷰 값 변경
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtEdit.value
                'DB 테이블 변경
                UpdateOrderListValue .selectedItem.Text, headerIndex, Me.txtEdit.value
                
                '수량,단가 변경한 경우에는 금액도 변경해야 함
                If headerIndex = 9 Or headerIndex = 11 Then
                    orderPrice = CalculateOrderListPrice(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(orderPrice, "#,##0")
                    UpdateOrderListValue .selectedItem.Text, 12, orderPrice
                End If
                '실행가 총액 계산
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
            Case 4  '분류
                fieldNo = 7
            Case 5  '거래처
                fieldNo = 8
            Case 6  '품목
                fieldNo = 9
            Case 7  '재질
                fieldNo = 10
            Case 8  '규격
                fieldNo = 11
            Case 9  '수량
                fieldNo = 12
            Case 10  '단위
                fieldNo = 13
            Case 11  '단가
                fieldNo = 14
            Case 12  '금액
                fieldNo = 15
            Case 13  '발주
                fieldNo = 18
            Case 14  '납기
                fieldNo = 19
            Case 15  '입고
                fieldNo = 20
            Case 16  '명세서
                fieldNo = 22
            Case 17  '계산서
                fieldNo = 23
            Case 18  '결제
                fieldNo = 24
        End Select
        
        shtOrderAdmin.Cells(findRow, fieldNo).value = value
    End If
End Sub

Private Sub btnOrderListInsert_Click()
    Dim lastId As Long
    Dim li As ListItem
    
    '발주리스트뷰에 발주 추가
    Insert_Record shtOrder, _
                , , "발주", currentManagementId, , , , , , , , , , _
                , , , , , _
                , , , , _
                , , _
                Date, , currentEstimateId, , False
    lastId = Get_LastID(shtOrder)
    
    With Me.lswOrderList
        Set li = .ListItems.Add(, , lastId)   'ID
        li.ListSubItems.Add , , currentEstimateId       'ID_견적
        li.ListSubItems.Add , , currentManagementId        '관리번호
        li.ListSubItems.Add , , "발주"        '분류
        li.ListSubItems.Add , , ""        '거래처
        li.ListSubItems.Add , , ""        '품목
        li.ListSubItems.Add , , ""        '재질
        li.ListSubItems.Add , , ""        '규격
        li.ListSubItems.Add , , ""        '수량
        li.ListSubItems.Add , , ""       '단위
        li.ListSubItems.Add , , ""          '단가
        li.ListSubItems.Add , , ""      '금액
        li.ListSubItems.Add , , ""       '발주일
        li.ListSubItems.Add , , ""       '납기일
        li.ListSubItems.Add , , ""       '입고일
        li.ListSubItems.Add , , ""       '명세서
        li.ListSubItems.Add , , ""       '계산서
        li.ListSubItems.Add , , ""       '결제
        li.ListSubItems.Add , , "열기"       '수정
        
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
        Case 4  '분류
            fieldName = "분류2"
        Case 5  '거래처
            fieldName = "거래처"
        Case 6  '품목
            fieldName = "품목"
        Case 7  '재질
            fieldName = "재질"
        Case 8  '규격
            fieldName = "규격"
        Case 9  '수량
            fieldName = "수량"
        Case 10  '단위
            fieldName = "단위"
        Case 11  '단가
            fieldName = "단가"
        Case 12  '금액
            fieldName = "금액"
        Case 13  '발주
            fieldName = "발주"
        Case 14  '납기
            fieldName = "납기"
        Case 15  '입고
            fieldName = "입고"
        Case 16  '명세서
            fieldName = "명세서"
        Case 17  '계산서
            fieldName = "계산서"
        Case 18  '결제
            fieldName = "결제"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtOrder, id, fieldName, value
        Update_Record_Column shtOrder, id, "수정일자", Date
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


