VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "견적 수정"
   ClientHeight    =   13110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19320
   OleObjectBlob   =   "frmEstimateUpdate.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEstimateUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim orgManagementID As Variant
Dim totlalCheckCount As Long
Dim mouseX As Integer
Dim headerIndex As Integer
Dim beforeSelectedItem As ListItem
Dim bInitialIzed As Boolean
Dim currentEditText, currentCboText As Variant




Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim db As Variant
    Dim contr As Control
    Dim acceptedMemo As Variant
    
    If clickEstimateId <> "" Then              '발주관리에서 더블클릭한 경우
        currentEstimateId = CLng(clickEstimateId)
        clickEstimateId = ""
    Else
        '선택한 행 번호
        cRow = Selection.row
    
        '데이터가 있는 행이 아닐 경우는 중지
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).value = "" Then
            MsgBox "수정할 견적 행을 먼저 선택한 후 견적수정 버튼을 클릭하세요.", vbInformation, "작업 확인"
            End
        End If
        
        currentEstimateId = shtEstimateAdmin.Cells(cRow, 3)
    End If
    
    bInitialIzed = False
    
     '텍스트박스 라벨 위치 조정
    For Each contr In Me.Controls
        If contr.Name Like "Label*" Then
            contr.top = contr.top + 2
        End If
    Next
    
    '폼 위치 수정
    If estimateUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateUpdateFormX
        Me.top = estimateUpdateFormY
    End If
    
    '견적 데이터 읽어오기
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If isEmpty(estimate) Or estimate(1) = "" Then
        MsgBox "견적 정보를 읽어올 수 없습니다. 견적ID(" & currentEstimateId & ")", vbInformation, "작업 확인"
        End
    End If

    Me.txtID.value = estimate(1)    'ID
    Me.txtEstimateName.value = estimate(6)  '견적명
    Me.txtManagementID.value = estimate(2)    '관리번호
    currentManagementId = Me.txtManagementID.value
    Me.txtLinkedID.value = estimate(3)  '자재번호
    
    Me.txtCustomer = estimate(4)   '거래처
    Me.txtManager = estimate(5)   '담당자
    
    Me.txtSize.value = estimate(7)  '규격
    Me.txtAmount.value = Format(estimate(8), "#,##0")   '수량
    InitializeCboUnit
    Me.cboUnit.value = Trim(estimate(9))  '단위, ID가 없으므로 직접 value 넣으면 선택됨
    Me.txtUnitPrice.value = Format(estimate(10), "#,##0")     '단가
    Me.txtEstimatePrice.value = Format(estimate(11), "#,##0")     '금액
    
    Me.txtEstimateDate.value = estimate(12)    '견적일자
    Me.txtBidDate.value = estimate(13)    '입찰일자
    Me.txtAcceptedDate.value = estimate(14)    '수주일자
    Me.txtDeliveryDate.value = estimate(15)    '납품일자
    Me.txtInsuranceDate.value = estimate(16)    '증권일자
    
    Me.txtProductionTotalCost.value = Format(estimate(17), "#,##0")   '(예상)실행가
    Me.txtBidPrice.value = Format(estimate(18), "#,##0")    '입찰가
    Me.txtBidMargin.value = Format(estimate(19), "#,##0")    '(예상)차액
    Me.txtBidMarginRate.value = Format(estimate(20), "0.0%")    '(예상)마진율
    Me.txtAcceptedPrice.value = Format(estimate(21), "#,##0")    '수주금액
    Me.txtAcceptedMargin.value = Format(estimate(22), "#,##0")   '수주차액
    If IsNumeric(Me.txtAcceptedPrice.value) And IsNumeric(Me.txtAcceptedMargin.value) And Me.txtAcceptedPrice <> "" Then
        If CLng(Me.txtAcceptedMargin.value) <> 0 Then
            If CLng(Me.txtAcceptedPrice.value) = 0 Then
                Me.txtAcceptedMarginRate = "0%"
            Else
                Me.txtAcceptedMarginRate = Format(CLng(Me.txtAcceptedMargin) / CLng(Me.txtAcceptedPrice.value), "0.0%")
            End If
        End If
    End If
    
    Me.txtInsertDate.value = estimate(23)    '등록일자
    Me.txtUpdateDate.value = estimate(24)    '수정일자
    
    InitializeCboCategory
    Me.cboCategory.value = Trim(estimate(25))   '분류1
    Me.txtDueDate.value = estimate(26)              '납기일
    
    Me.txtVAT.value = Format(estimate(31), "#,##0")    '부가세
    Me.txtMemo.value = Trim(estimate(32))     '견적메모
    Me.chkVAT.value = estimate(33)      '부가세 제외 여부
    
    Me.txtPaid.value = Format(estimate(34), "#,##0")      '입금액
    Me.txtRemaining.value = Format(estimate(35), "#,##0")      '미입금액
    
    '예전에는 견적관리xls와 관리xls에서 메뉴가 다르게 남음
    '앞으로 견적메모와 수주메모는 같게 맞춰야 함.
    '견적메모와 수주메모가 다른 경우는 예전 경우임
    '(견적메모 = 견적메모 + 수주메모) 이렇게 맞추고 저장 시 수주쪽에도 동일하게 메모 넣을 예정
    acceptedMemo = Trim(estimate(36))
    If Me.txtMemo.value <> acceptedMemo Then
        If Me.txtMemo.value = "" Then
            Me.txtMemo.value = acceptedMemo
        Else
            Me.txtMemo.value = Me.txtMemo.value & vbCrLf & acceptedMemo
        End If
    End If
    
    '수주 ID (ID_관리)
    Me.txtAcceptedID.value = estimate(37)
    currentAcceptedId = estimate(37)
    If Me.txtAcceptedID.value = "" Then
        '수주ID가 없으면 수주관련 컨트롤 unable 시킴
        frmOrder.Visible = False
        btnAcceptedInsert.Visible = True
        frmEstimateUpdate.Height = 260
    Else
        frmOrder.Visible = True
        btnAcceptedInsert.Visible = False
    End If
    
    '변경 전 관리번호
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '발주 현황
    InitializeCboOrderCategory  '발주 분류2
    InitializeLswCustomerAutoComplete   '거래처 자동완성
    InitializeLswManagerAutoComplete    '담당자 자동완성
    InitializeLswPaymentList      '결제 현황
    InitializeCboEstimatePayMethod '결제수단-수입
    
    Me.txtSize.SetFocus
    
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtEstimateCustomer, True)

    Update_Cbo Me.cboCustomer, db
End Sub


Sub InitializeCboManager()
    Dim db As Variant
    Dim i As Long
    
    '담당자 DB를 읽어와서
    db = Get_DB(shtEstimateManager, True)
    '거래처명으로 필터링
    db = Filtered_DB(db, Me.cboCustomer.value, 1, True)
    
    '기존 콤보박스 내용지우기
    Me.cboManager.Clear
    
    '담당자가 있으면 콤보박스에 추가함
    If Not isEmpty(db) Then
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

Sub InitializeCboOrderCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboOrderCategory, db
End Sub

Sub InitializeCboEstimatePayMethod()
    Dim db As Variant
    db = Get_DB(shtEstimatePayMethod, True)

    Update_Cbo Me.cboEstimatePayMethod, db
End Sub

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j As Long
    Dim totalCost As Double
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

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeLswManagerAutoComplete()
    
    With Me.lswManagerAutoComplete
        .View = lvwList
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeLswPaymentList()
    Dim db As Variant
    Dim i, j, totalPaid As Long
    Dim li As ListItem

    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With

     '리스트뷰 값 설정
    With Me.lswPaymentList
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
        .ColumnHeaders.Add , , "명세서", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "계산서", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "결제", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "결제월", 59, lvwColumnCenter
        .ColumnHeaders.Add , , "입금액", 70, lvwColumnRight
        .ColumnHeaders.Add , , "입금예정액", 70, lvwColumnRight
        .ColumnHeaders.Add , , "결제수단", 60, lvwColumnCenter
        .ColumnHeaders.Add , , "메모", 110
        .ColumnHeaders.Add , , "부가세", 60, lvwColumnRight
        .ColumnHeaders.Add , , "결제예정일", 0

        .ListItems.Clear
        
        '견적ID에 해당하는 결제 이력를 읽어옴
        db = Get_DB(shtPayment)
        If Not isEmpty(db) Then
            db = Filtered_DB(db, Me.txtID.value, 2, True)
        End If

        totalPaid = 0
        If Not isEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))   'ID
                li.ListSubItems.Add , , db(i, 2)       'ID_견적
                li.ListSubItems.Add , , db(i, 3)        '관리번호
                li.ListSubItems.Add , , db(i, 4)        '명세서
                li.ListSubItems.Add , , db(i, 5)        '계산서
                li.ListSubItems.Add , , db(i, 6)        '결제
                li.ListSubItems.Add , , Format(db(i, 7), "mm" & "월")      '결제월
                li.ListSubItems.Add , , Format(db(i, 8), "#,##0")        '입금액
                li.ListSubItems.Add , , Format(db(i, 9), "#,##0")        '입금예정액
                li.ListSubItems.Add , , db(i, 10)        '결제수단
                li.ListSubItems.Add , , db(i, 11)       '메모
                li.ListSubItems.Add , , Format(db(i, 12), "#,##0")       '부가세
                li.ListSubItems.Add , , db(i, 7)       '결제예정일
                li.Selected = False

                If IsNumeric(db(i, 8)) Then
                    '입금 합계 구함
                    totalPaid = totalPaid + CLng(db(i, 8))
                End If
            Next
        End If

        Me.txtPaid.value = Format(totalPaid, "#,##0")
        CalculatePayment
        
    End With

End Sub

Sub UpdateEstimateOrderValue(fieldName, fieldValue)
        
    '견적DB 변경
    Update_Record_Column shtEstimate, currentEstimateId, fieldName, fieldValue
    
    '견적시트 변경
    UpdateShtEstimateField currentEstimateId, fieldName, fieldValue
    
    '수주DB 변경
    Update_Record_Column shtOrder, currentAcceptedId, fieldName, fieldValue
    
    '수주시트 변경
    UpdateShtOrderField currentAcceptedId, fieldName, fieldValue
    
End Sub

Sub UpdateEstimateValue(fieldName, fieldValue)
        
    '견적DB 변경
    Update_Record_Column shtEstimate, currentEstimateId, fieldName, fieldValue
    
    '견적시트 변경
    UpdateShtEstimateField currentEstimateId, fieldName, fieldValue
End Sub

Sub UpdateOrderValue(fieldName, fieldValue)
    '수주DB 변경
    Update_Record_Column shtOrder, currentAcceptedId, fieldName, fieldValue
    
    '수주시트 변경
    UpdateShtOrderField currentAcceptedId, fieldName, fieldValue
    
End Sub

Sub InsertAccepted()

    '수주발주 테이블에 수주 등록
    Insert_Record shtOrder, _
            , Me.cboCategory.value, "수주", Me.txtManagementID.value, _
            Me.txtCustomer.value, _
            Me.txtEstimateName.value, _
            Me.txtManager.value, _
            Me.txtSize.value, _
            Me.txtAmount.value, _
            Me.cboUnit.value, _
            , _
            , _
            , _
            Date, , , , , _
            , , , , _
            , , _
            Date, , _
            CLng(Me.txtID.value), , False, "수입"

    '등록한 수주ID를 견적 테이블에 업데이트, 수주일자는 오늘
    Update_Record_Column shtEstimate, Me.txtID, "ID_수주", Get_LastID(shtOrder)
    Update_Record_Column shtEstimate, Me.txtID, "수주", Date
    
    '폼을 새로 띄움
    Unload frmEstimateUpdate
    
    clickEstimateId = Me.txtID
    frmEstimateUpdate.Show (False)
    
End Sub

'발주 리스트뷰 값 DB 저장
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
        UpdateShtOrderField id, fieldName, value
    End If

End Sub

Sub UpdatePaymentListValue(id, headerIndex, value)
    Dim fieldName As String

    Select Case headerIndex
        Case 4
            fieldName = "명세서"
        Case 5
            fieldName = "계산서"
        Case 6
            fieldName = "결제"
        Case 7
            fieldName = "결제월"
        Case 8
            fieldName = "입금액"
        Case 9
            fieldName = "입금예정액"
        Case 10
            fieldName = "결제수단"
        Case 11
            fieldName = "메모"
        Case 12
            fieldName = "부가세"
    End Select
    
    If fieldName <> "" Then
        Update_Record_Column shtPayment, id, fieldName, value
        Update_Record_Column shtPayment, id, "수정일자", Date
    End If
    
    Select Case headerIndex
        Case 4, 5, 6, 7, 10, 12
            '만약 현재 선택한 행이 맨 마지막행이면 명세서/계산서/결제/결제월 데이터를 견적DB/발주DB/견적관리시트/발주관리시트에 저장
            If lswPaymentList.ListItems(lswPaymentList.ListItems.count).Selected = True Then
                UpdateEstimateOrderValue fieldName, value
            End If
    End Select

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

Sub SelectPaymentListColumn()
    Dim ItemSel    As ListItem
    
    If Not lswPaymentList.selectedItem Is Nothing Then
        If headerIndex = lswPaymentList.ColumnHeaders.count Then
            frmPaymentEdit.Visible = False
            txtPaymentEdit.Visible = False
            cboEstimatePayMethod.Visible = False
        End If
        
        Set ItemSel = lswPaymentList.selectedItem
        ItemSel.EnsureVisible
        
        If headerIndex = 10 Then
            With frmPaymentEdit
                .Visible = True
                .top = ItemSel.top + lswPaymentList.top
                .Left = lswPaymentList.ColumnHeaders(headerIndex).Left + lswPaymentList.Left
                .Width = Me.cboEstimatePayMethod.Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.cboEstimatePayMethod
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Height = lswPaymentList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentCboText = .Text
            End With
            Me.txtPaymentEdit.Visible = False
        ElseIf headerIndex > 3 And headerIndex < lswPaymentList.ColumnHeaders.count Then
            With frmPaymentEdit
                .Visible = True
                .top = ItemSel.top + lswPaymentList.top
                .Left = lswPaymentList.ColumnHeaders(headerIndex).Left + lswPaymentList.Left
                .Width = lswPaymentList.ColumnHeaders(headerIndex).Width
                .Height = ItemSel.Height + 2
                .ZOrder msoBringToFront
            End With
            With Me.txtPaymentEdit
                .Visible = True
                .Text = ItemSel.SubItems(headerIndex - 1)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .top = 0
                .Width = lswPaymentList.ColumnHeaders(headerIndex).Width
                .Height = lswPaymentList.selectedItem.Height + 2
                .SelLength = Len(.Text)
                currentEditText = .Text
            End With
            Me.cboEstimatePayMethod.Visible = False
        End If
    End If

End Sub

Sub DeleteOrderList()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "삭제할 발주를 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    yn = MsgBox("선택한 " & count & "개 발주를 삭제할까요?", vbYesNo + vbQuestion, "작업 확인")
    If yn = vbNo Then Exit Sub

    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            '발주 테이블에서 삭제
            Delete_Record shtOrder, li.Text
        End If
    Next
    
    If count > 0 Then
        InitializeLswOrderList
    End If
End Sub
'
'Sub UpdateShtEstimate(estimateId)
'    Dim findRow As Long
'
'    findRow = isExistInSheet(shtEstimateAdmin.Range("B6"), estimateId)
'    If findRow <> 0 Then
'        shtEstimateAdmin.Cells(findRow, 4).value = Me.txtManagementID.value
'        shtEstimateAdmin.Cells(findRow, 5).value = Me.txtCustomer.value
'        shtEstimateAdmin.Cells(findRow, 6).value = Me.txtManager.value
'        shtEstimateAdmin.Cells(findRow, 7).value = Me.cboCategory.value
'        shtEstimateAdmin.Cells(findRow, 8).value = Me.txtEstimateName.value
'        shtEstimateAdmin.Cells(findRow, 9).value = Me.txtSize.value
'        shtEstimateAdmin.Cells(findRow, 10).value = Me.txtAmount.value
'        shtEstimateAdmin.Cells(findRow, 11).value = Me.cboUnit.value
'        shtEstimateAdmin.Cells(findRow, 12).value = Me.txtUnitPrice.value
'        shtEstimateAdmin.Cells(findRow, 13).value = Me.txtEstimatePrice.value
'        shtEstimateAdmin.Cells(findRow, 14).value = Me.txtEstimateDate.value
'        shtEstimateAdmin.Cells(findRow, 15).value = Me.txtBidDate.value
'        shtEstimateAdmin.Cells(findRow, 16).value = Me.txtAcceptedDate.value
'        shtEstimateAdmin.Cells(findRow, 17).value = Me.txtDueDate.value
'        shtEstimateAdmin.Cells(findRow, 18).value = Me.txtDeliveryDate.value
'        shtEstimateAdmin.Cells(findRow, 19).value = Me.txtInsuranceDate.value
'        shtEstimateAdmin.Cells(findRow, 20).value = Me.txtProductionTotalCost.value
'        shtEstimateAdmin.Cells(findRow, 21).value = Me.txtBidPrice.value
'        shtEstimateAdmin.Cells(findRow, 22).value = Me.txtBidMargin.value
'        shtEstimateAdmin.Cells(findRow, 23).value = Me.txtBidMarginRate.value
'        shtEstimateAdmin.Cells(findRow, 24).value = Me.txtAcceptedPrice.value
'        shtEstimateAdmin.Cells(findRow, 25).value = Me.txtAcceptedMargin.value
'        shtEstimateAdmin.Cells(findRow, 26).value = Me.txtSpecificationDate.value
'        shtEstimateAdmin.Cells(findRow, 27).value = Me.txtTaxinvoiceDate.value
'        shtEstimateAdmin.Cells(findRow, 28).value = Me.txtPaymentDate.value
'        shtEstimateAdmin.Cells(findRow, 29).value = Me.txtExpectPaymentDate.value
'        shtEstimateAdmin.Cells(findRow, 30).value = Me.txtVAT.value
'        shtEstimateAdmin.Cells(findRow, 31).value = Me.txtInsertDate.value
'        shtEstimateAdmin.Cells(findRow, 32).value = Date
'    End If
'End Sub

Sub UpdateShtEstimateField(estimateId, fieldName, value)
    Dim findRow As Long
    Dim colNo As Long
    
    findRow = isExistInSheet(shtEstimateAdmin.Range("C6"), estimateId)
    If findRow > 0 Then
        colNo = 0
        Select Case fieldName
            Case "관리번호"
                colNo = 4
            Case "거래처"
                colNo = 5
            Case "담당자"
                colNo = 6
            Case "분류1"
                colNo = 7
            Case "견적명"
                colNo = 8
            Case "규격"
                colNo = 9
            Case "수량"
                colNo = 10
            Case "단위"
                colNo = 11
            Case "단가"
                colNo = 12
            Case "금액"
                colNo = 13
            Case "견적"
                colNo = 14
            Case "입찰"
                colNo = 15
            Case "수주"
                colNo = 16
            Case "납기"
                colNo = 17
            Case "납품"
                colNo = 18
            Case "증권"
                colNo = 19
            Case "실행가(예상)"
                colNo = 20
            Case "입찰금액"
                colNo = 21
            Case "차액(예상)"
                colNo = 22
            Case "마진율(예상)"
                colNo = 23
            Case "수주금액"
                colNo = 24
            Case "수주차액"
                colNo = 25
            Case "명세서"
                colNo = 26
            Case "계산서"
                colNo = 27
            Case "결제"
                colNo = 28
            Case "결제월"
                colNo = 29
            Case "부가세"
                colNo = 30
            Case "등록일자"
                colNo = 31
            Case "수정일자"
                colNo = 32
        End Select
      
        If colNo <> 0 Then
            shtEstimateAdmin.Cells(findRow, colNo).value = value
        End If
    End If
    
End Sub

'Sub UpdateShtOrder(orderId)
'    Dim findRow As Long
'
'    findRow = isExistInSheet(shtOrderAdmin.Range("C6"), orderId)
'    If findRow <> 0 Then
'        shtOrderAdmin.Cells(findRow, 5).value = Me.txtManagementID.value
'        shtOrderAdmin.Cells(findRow, 6).value = Me.cboCategory.value
'        shtOrderAdmin.Cells(findRow, 8).value = Me.txtCustomer.value
'        shtOrderAdmin.Cells(findRow, 9).value = Me.txtEstimateName.value
'        shtOrderAdmin.Cells(findRow, 10).value = Me.txtManager.value
'        shtOrderAdmin.Cells(findRow, 11).value = Me.txtSize.value
'        shtOrderAdmin.Cells(findRow, 12).value = Me.txtAmount.value
'        shtOrderAdmin.Cells(findRow, 13).value = Me.cboUnit.value
'        shtOrderAdmin.Cells(findRow, 14).value = Me.txtUnitPrice.value
'        shtOrderAdmin.Cells(findRow, 15).value = Me.txtEstimatePrice.value
'        shtOrderAdmin.Cells(findRow, 17).value = Me.txtAcceptedDate.value
'        shtOrderAdmin.Cells(findRow, 19).value = Me.txtDueDate.value
'        shtOrderAdmin.Cells(findRow, 21).value = Me.txtDeliveryDate.value
'        shtOrderAdmin.Cells(findRow, 22).value = Me.txtSpecificationDate.value
'        shtOrderAdmin.Cells(findRow, 23).value = Me.txtTaxinvoiceDate.value
'        shtOrderAdmin.Cells(findRow, 24).value = Me.txtPaymentDate.value
'        shtOrderAdmin.Cells(findRow, 25).value = Me.txtExpectPaymentDate.value
'        shtOrderAdmin.Cells(findRow, 27).value = Me.txtVAT.value
'        shtOrderAdmin.Cells(findRow, 28).value = Me.txtInsertDate.value
'        shtOrderAdmin.Cells(findRow, 29).value = Date
'    End If
'End Sub

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

Function CheckEstimateUpdateValidation()
    
    '견적명이 입력되었는지 체크
    If Me.txtEstimateName.value = "" Then
        MsgBox "견적명을 입력하세요.", vbExclamation
        CheckEstimateUpdateValidation = False
        Exit Function
    End If
    
    '관리번호가 입력되었는지 체크
    If Me.txtManagementID.value = "" Then
        MsgBox "관리번호를 입력하세요.", vbExclamation
        CheckEstimateUpdateValidation = False
        Exit Function
    End If
    
    CheckEstimateUpdateValidation = True
End Function


Sub CalculateEstimatePrice()
    '금액 계산
    '수량값이 공백이면 금액은 단가
    If Me.txtUnitPrice = "" Then
        Me.txtEstimatePrice.value = ""
    Else
        If Me.txtAmount.value = "" Then
            Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Else
            Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")
End Sub

Sub CalculateBidMargin()
    '예상차액과 예상마진율 계산
    If Me.txtBidPrice.value <> "" And Me.txtProductionTotalCost.value <> "" Then
        '예상차액 = 입찰가 - 예상실행가
        Me.txtBidMargin.value = Format(CLng(Me.txtBidPrice.value) - CLng(Me.txtProductionTotalCost.value), "#,##0")
        '예상마진율 = 예상차액 / 입찰가
        If Me.txtBidPrice.value <> "0" Then
            Me.txtBidMarginRate.value = Format(CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value), "0.0%")
        End If
    Else
        Me.txtBidMargin.value = ""
        Me.txtBidMarginRate.value = ""
    End If
End Sub

Sub CalculateAcceptedMargin()

    '수주차액, 마진율 계산
    If Me.txtAcceptedPrice.value <> "" And Me.txtExecutionCost.value <> "" Then
        '수주차액 = 수주금액 - 실행가
        Me.txtAcceptedMargin.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtExecutionCost.value), "#,##0")
        '마진율 = 수주차액 / 수주금액
        If Me.txtAcceptedPrice.value <> "0" Then
            Me.txtAcceptedMarginRate.value = Format(CLng(Me.txtAcceptedMargin.value) / CLng(Me.txtAcceptedPrice.value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.value = ""
        Me.txtAcceptedMarginRate.value = ""
    End If
End Sub

Sub CalculatePayment()
   '미입금액, 부가세 계산
    If Me.txtAcceptedPrice.value = "" Then
        Me.txtRemaining.value = ""
        Me.txtVAT.value = ""
    Else
        If IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtRemaining.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtPaid.value), "#,##0")
            If Me.chkVAT.value = True Then
                Me.txtVAT.value = 0
            Else
                Me.txtVAT = Format(CLng(Me.txtPaid.value) * 0.1, "#,##0")
            End If
        End If
    End If
End Sub

Sub CalculateEstimateUpdateCost_2()

    '금액 계산
    '수량값이 공백이면 금액은 단가
    If Me.txtUnitPrice = "" Then
        Me.txtEstimatePrice.value = ""
    Else
        If Me.txtAmount.value = "" Then
            Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Else
            Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")

    '예상차액과 예상마진율 계산
    If Me.txtBidPrice.value <> "" And Me.txtProductionTotalCost.value <> "" Then
        '예상차액 = 입찰가 - 예상실행가
        Me.txtBidMargin.value = Format(CLng(Me.txtBidPrice.value) - CLng(Me.txtProductionTotalCost.value), "#,##0")
        '예상마진율 = 예상차액 / 입찰가
        If Me.txtBidPrice.value <> "0" Then
            Me.txtBidMarginRate.value = Format(CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value), "0.0%")
        End If
    Else
        Me.txtBidMargin.value = ""
        Me.txtBidMarginRate.value = ""
    End If

    '수주차액, 마진율 계산
    If Me.txtAcceptedPrice.value <> "" And Me.txtExecutionCost.value <> "" Then
        '수주차액 = 수주금액 - 실행가
        Me.txtAcceptedMargin.value = Format(CLng(Me.txtAcceptedPrice.value) - CLng(Me.txtExecutionCost.value), "#,##0")
        '마진율 = 수주차액 / 수주금액
        If Me.txtAcceptedPrice.value <> "0" Then
            Me.txtAcceptedMarginRate.value = Format(CLng(Me.txtAcceptedMargin.value) / CLng(Me.txtAcceptedPrice.value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.value = ""
        Me.txtAcceptedMarginRate.value = ""
    End If

    '부가세 계산
    '세금계산서 일자가 없는 경우, 부가세 제외인 경우 부가세는 0
    If Me.txtTaxinvoiceDate.value = "" Or chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        '부가세는 수주금액의 10%
        If Me.txtAcceptedPrice.value <> "" And Me.txtAcceptedPrice.value <> 0 Then
            Me.txtVAT.value = CLng(Me.txtAcceptedPrice.value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.value, "#,##0")
        End If
    End If

'    '입금예상액 계산
'    If Me.txtTaxInvoiceDate.Value = "" Then
'        '세금계산서 일자가 없는 경우는 수주금액
'        Me.txtExpectPay.Value = Me.txtAcceptedPrice
'    Else
'        '세금계산서 일자가 있는 경우는 수주금액+부가세
'        If Me.txtAcceptedPrice.Value <> "" Then
'            Me.txtExpectPay.Value = CLng(Me.txtAcceptedPrice.Value) + CLng(Me.txtVAT.Value)
'        End If
'    End If
'    Me.txtExpectPay.Text = Format(Me.txtExpectPay.Value, "#,##0")
'
'    '입금액 계산
'    If Me.txtPaymentDate.Value = "" Then
'        Me.txtPaid.Value = 0
'    Else
'        Me.txtPaid.Value = Me.txtExpectPay.Value
'        Me.txtPaid.Text = Format(Me.txtPaid.Value, "#,##0")
'    End If
'
'    '미입금액 계산
'    Me.txtUnpaid.Value = CLng(Me.txtExpectPay.Value) - CLng(Me.txtPaid.Value)
'    Me.txtUnpaid.Text = Format(Me.txtUnpaid.Value, "#,##0")
    
End Sub

Function CalculateOrderListTotalCost() As Long
    Dim i As Long
    Dim cost, totalCost As Long
    
    With Me.lswOrderList
        For i = 1 To .ListItems.count

            If Not IsNumeric(.ListItems(i).SubItems(11)) Then
                If .ListItems(i).SubItems(11) <> "" Then
                    MsgBox "금액 필드에 숫자가 아닌 값이 있어서 실행가 합계를 구할 수 없습니다.", vbExclamation
                    CalculateOrderListTotalCost = 0
                    Exit Function
                End If
            Else
                totalCost = totalCost + .ListItems(i).SubItems(11)
            End If
        Next
    End With
    
    CalculateOrderListTotalCost = totalCost
End Function

Function CalculatePaymentListTotalCost() As Long
    Dim i As Long
    Dim cost, totalCost As Long
    
    With Me.lswPaymentList
        For i = 1 To .ListItems.count

            If Not IsNumeric(.ListItems(i).SubItems(7)) Then
                If .ListItems(i).SubItems(7) <> "" Then
                    MsgBox "금액 필드에 숫자가 아닌 값이 있어서 입금액 합계를 구할 수 없습니다.", vbExclamation
                    CalculatePaymentListTotalCost = 0
                    Exit Function
                End If
            Else
                totalCost = totalCost + .ListItems(i).SubItems(7)
            End If
        Next
    End With
    
    CalculatePaymentListTotalCost = totalCost
End Function

Function CalculateOrderListPrice(selectedItem As ListItem) As Long
    Dim amount, unitPrice As Variant
    Dim orderPrice As Long

    '수량, 단가가 변하는 경우에는 금액 계산해서 변경해야 함
    amount = selectedItem.ListSubItems(8).Text
    unitPrice = selectedItem.ListSubItems(10).Text
    
    If amount = "" Then
        If IsNumeric(unitPrice) Then
            orderPrice = unitPrice
        End If
    ElseIf IsNumeric(amount) And IsNumeric(unitPrice) Then
        orderPrice = amount * unitPrice
    End If
    
    CalculateOrderListPrice = orderPrice
End Function

Function CalculatePaymentListVAT(selectedItem As ListItem) As Long
    Dim paid As Variant
    Dim VAT As Long

    '계산서 값이 없으면 0
    If selectedItem.ListSubItems(4).Text = "" Then
        CalculatePaymentListVAT = 0
        Exit Function
    End If
    
    '입금액 변하는 경우에 부가세 변경해야 함
    paid = selectedItem.ListSubItems(7).Text
    
    If paid = "" Then
        '입금액이 없는 경우에 입금예정액으로 함
        paid = selectedItem.ListSubItems(8).Text
    End If
    
    If IsNumeric(paid) Then
        VAT = paid * 0.1
    Else
        VAT = 0
    End If
    
    CalculatePaymentListVAT = VAT
End Function

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.value, 2, True)
    
    'DB에 값이 있을 경우
    totalCost = 0
    If Not isEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 11)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(db(i, 11))
            End If
        Next
    End If
        
    GetProductionTotalCost = totalCost
End Function

Sub ConvertOrderListFormat(textBox, headerIndex)
    Dim value As Variant
    Dim pos As Long
    Dim Y, M, D As Long
    
    value = Trim(textBox.Text)
    
    Select Case headerIndex
        Case 9, 11, 12  '수량, 단가, 금액 - 1000자리 콤마
            If IsNumeric(value) Then
                textBox.Text = Format(value, "#,##0")
            End If
        Case 13, 14, 15, 16, 17, 18  '발주, 납기, 입고, 명세서, 계산서, 결제 - 날짜 변환
            textBox.Text = ConvertDateFormat(textBox.Text)
    End Select
    
End Sub

Sub ConvertPaymentListFormat(textBox, headerIndex)
    Select Case headerIndex
        Case 8, 9, 12  '입금액, 부가세 - 1000자리 콤마
            If IsNumeric(textBox.Text) Then
                textBox.Text = Format(textBox.Text, "#,##0")
            End If
        Case 4, 5, 6, 7  '명세서, 계산서, 결제, 결제월 - 날짜 변환
            textBox.Text = ConvertDateFormat(textBox.Text)
    End Select
    
End Sub




Private Sub lswOrderList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * x
End Sub

Private Sub lswPaymentList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    mouseX = pointsPerPixelX * x
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

Private Sub btnOrderListBatchUpdate_Click()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "일괄 변경할 발주를 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    frmOrderDateUpdate.Show
End Sub

Private Sub btnPaymentListInsert_Click()
    Dim lastId As Long
    Dim li As ListItem
    
    frmPaymentEdit.Visible = False
    cboEstimatePayMethod.Visible = False
    
    '결제이력에 발주 추가
    Insert_Record shtPayment, _
                        currentEstimateId, currentManagementId, _
                        Date, , , , , , , , , Date, ""
                        
    lastId = Get_LastID(shtPayment)
    
    With Me.lswPaymentList
        Set li = .ListItems.Add(, , lastId)   'ID
        li.ListSubItems.Add , , currentEstimateId       'ID_견적
        li.ListSubItems.Add , , currentManagementId        '관리번호
        li.ListSubItems.Add , , Date        '명세서
        li.ListSubItems.Add , , ""        '계산서
        li.ListSubItems.Add , , ""        '결제
        li.ListSubItems.Add , , ""       '결제월
        li.ListSubItems.Add , , ""        '입금액
        li.ListSubItems.Add , , ""        '입금예정액
        li.ListSubItems.Add , , ""       '결제수단
        li.ListSubItems.Add , , ""       '메모
        li.ListSubItems.Add , , ""      '부가세
        li.ListSubItems.Add , , ""      '결제예정일(날짜포맷)
        
        .selectedItem.Selected = False
        li.Selected = True
        li.EnsureVisible
        
        headerIndex = 4
        SelectPaymentListColumn
    End With
    
    'DB와 시트의 명세서/계산서/결제/결제월 값 변경
    UpdateEstimateOrderValue "명세서", Date
    UpdateEstimateOrderValue "계산서", ""
    UpdateEstimateOrderValue "결제", ""
    UpdateEstimateOrderValue "결제월", ""
    UpdateEstimateOrderValue "결제수단", ""
    
End Sub

Private Sub btnPaymentListDelete_Click()
    Dim li As ListItem
    Dim count As Long
    Dim yn As VbMsgBoxResult
    Dim spec, tax, paid, month, method As Variant
    
    count = 0
    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "삭제할 결제이력을 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    yn = MsgBox("선택한 " & count & "개 이력을 삭제할까요?", vbYesNo + vbQuestion, "작업 확인")
    If yn = vbNo Then Exit Sub

    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            '이력 테이블에서 삭제
            Delete_Record shtPayment, li.Text
        End If
    Next
    
    If count > 0 Then
        InitializeLswPaymentList
    End If
    
    '삭제 플래그 세팅해서 결제이력이 추가되지 않도록 함
    bDeleteFlag = True
    
    '맨 마지막 이력 데이터의 명세서/계산서/결제/결제월을 견적DB에 저장
    With Me.lswPaymentList
        If .ListItems.count = 0 Then
            spec = ""
            tax = ""
            paid = ""
            month = ""
            method = ""
        Else
            spec = .ListItems(.ListItems.count).SubItems(3)
            tax = .ListItems(.ListItems.count).SubItems(4)
            paid = .ListItems(.ListItems.count).SubItems(5)
            month = .ListItems(.ListItems.count).SubItems(6)
            method = .ListItems(.ListItems.count).SubItems(8)
        End If
        UpdateEstimateOrderValue "명세서", spec
        UpdateEstimateOrderValue "계산서", tax
        UpdateEstimateOrderValue "결제", paid
        UpdateEstimateOrderValue "결제월", month
        UpdateEstimateOrderValue "결제수단", method
    End With
    
    bDeleteFlag = False
End Sub


Private Sub lswOrderList_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.value = ""
    Me.cboOrderCategory.value = ""
End Sub

Private Sub lswPaymentList_Click()
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.value = ""
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

Private Sub lswPaymentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswPaymentList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub btnProduction_Click()
    If isFormLoaded("frmProductionManager") Then
        Unload frmProductionManager
    End If
    frmProductionManager.Show (False)
End Sub


Private Sub btnAcceptedInsert_Click()
    InsertAccepted
End Sub

Private Sub btnPayment_Click()
    frmPaymentManager.Show
End Sub

Private Sub chkDividePay_Click()
    If chkDividePay.value = True Then
        Me.btnPayment.Enabled = True
    Else
        Me.btnPayment.Enabled = False
    End If
End Sub

Private Sub btnEstimateUpdate_Click()
    UpdateEstimate
End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnOrderListDelete_Click()
    DeleteOrderList
End Sub

Private Sub Frame4_Click()
    Me.frmEdit.Visible = False
    Me.txtEdit.Visible = False
    Me.cboOrderCategory.Visible = False
    Me.frmPaymentEdit.Visible = False
    Me.txtPaymentEdit.Visible = False
End Sub

Private Sub frmOrder_Click()
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

Private Sub lswCustomerAutoComplete_DblClick()
    '거래처에 값을 넣어주고 포커스는 매니저로 이동
    With Me.lswCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
        End If
    End With
End Sub

Private Sub lswManagerAutoComplete_DblClick()
    '담당자명에 값을 넣어주고 포커스는 견적명으로 이동
    With Me.lswManagerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtManager.value = .selectedItem.Text
            .Visible = False
            Me.txtEstimateName.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderList_DblClick()

    Dim i As Integer
    Dim pos As Integer
    
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

Private Sub lswPaymentList_DblClick()

    Dim i As Integer
    Dim pos As Integer
    
    With Me.lswPaymentList
        headerIndex = 0
        For i = 1 To .ColumnHeaders.count
            pos = .ColumnHeaders(i).Left
            If mouseX < pos Then
                headerIndex = i - 1
                Exit For
            End If
        Next
        
        If headerIndex = 12 Then
            '부가세는 변경할 수 없음
        ElseIf headerIndex >= 4 Then
            ' 현재 선택한 열을 저장해놓음
            If Not beforeSelectedItem Is Nothing Then
                Set beforeSelectedItem = Nothing
            End If
            Set beforeSelectedItem = .selectedItem
            
            SelectPaymentListColumn
        End If
    End With

End Sub

Private Sub txtManager_Enter()
    '자동완성 리스트에서 탭해서 넘어오는 경우
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtEstimateName_Enter()
    '자동완성 리스트에서 탭해서 넘어오는 경우
    With Me.lswManagerAutoComplete
        If .Visible = True Then
            Me.txtManager.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub lbl2AcceptedDate_Enter()
    Me.txtAcceptedDate.SetFocus
End Sub

Private Sub lbl2BidDate_Enter()
    Me.txtBidDate.SetFocus
End Sub

Private Sub lbl2DeliveryDate_Enter()
    Me.txtDeliveryDate.SetFocus
End Sub

Private Sub lbl2EstimateDate_Enter()
    Me.txtEstimateDate.SetFocus
End Sub

Private Sub UserForm_Activate()
    If bInitialIzed = False Then
        Me.txtSize.SetFocus
    End If
    bInitialIzed = True
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = vbKeyReturn Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtManager.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '탭키일 경우에 자동완성 결과가 하나이면 다음 입력칸으로 이동
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtManager.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = vbKeyDown Then
            '아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
End Sub

Private Sub txtCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '거래처 자동완성 처리
    With Me.lswCustomerAutoComplete
        If Me.txtCustomer.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '견적거래처 DB를 읽어와서 리스트뷰에 출력
            .ListItems.Clear
            db = Get_DB(shtEstimateCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.value, 1, False)
            If isEmpty(db) Then
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

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswManagerAutoComplete
        If KeyCode = vbKeyReturn Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtEstimateName.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            '탭키일 경우에 자동완성 결과가 하나이면 다음 입력칸으로 이동
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtEstimateName.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = vbKeyDown Then
            '아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
    
End Sub

Private Sub txtSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub chkDividePay_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.btnEstimateUpdate.SetFocus
    End If
End Sub


Private Sub txtManager_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '담당자 자동완성 처리
    With Me.lswManagerAutoComplete
        If Me.txtManager.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '견적담당자 DB를 읽어와서 리스트뷰에 출력
            .ListItems.Clear
            db = Get_DB(shtEstimateManager, True)
            db = Filtered_DB(db, Me.txtManager.value, 1, False)
            If isEmpty(db) Then
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

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처에 값을 넣어주고 포커스는 매니저로 이동
    If KeyCode = vbKeyReturn Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
End Sub

Private Sub lswManagerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '담당자 선택 후 엔터키 들어오면 이 값을 담당자명에 넣어주고 포커스는 다음(사이즈)으로 이동
    If KeyCode = vbKeyReturn Then
        With Me.lswManagerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtManager.value = .selectedItem.Text
                .Visible = False
                Me.txtEstimateName.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub btnProduction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub


Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyTab Then
'        MsgBox Me.txtMemo.CurTargetX & ", " & Me.txtMemo.CurX
'        KeyCode = 0
'    End If
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
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

Sub PaymentListUpdate(headerIndex)
    Dim VAT As Variant
    
    With Me.lswPaymentList
        If .selectedItem Is Nothing Then
            Exit Sub
        End If
        
        If headerIndex = 10 Then
            If Me.cboEstimatePayMethod.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '리스트뷰 값 변경
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.cboEstimatePayMethod.value
                'DB 테이블 변경
                UpdatePaymentListValue .selectedItem.Text, headerIndex, Me.cboEstimatePayMethod.value
            End If
        Else
            If Me.txtPaymentEdit.value <> .selectedItem.ListSubItems(headerIndex - 1).Text Then
                '입력값 포맷 변경
                ConvertPaymentListFormat Me.txtPaymentEdit, headerIndex
                '리스트뷰 값 변경
                .selectedItem.ListSubItems(headerIndex - 1).Text = Me.txtPaymentEdit.value
                'DB 테이블 변경
                UpdatePaymentListValue .selectedItem.Text, headerIndex, Me.txtPaymentEdit.value
                
                If headerIndex = 7 Then
                    '결제월인 경우에 화면에 '10월' 이런식으로 보여줌
                    .selectedItem.ListSubItems(6).Text = Format(Me.txtPaymentEdit.value, "mm" & "월")
                    .selectedItem.ListSubItems(12).Text = Me.txtPaymentEdit.value
                ElseIf headerIndex = 5 Or headerIndex = 8 Or headerIndex = 9 Then
                    '입금액 변경한 경우에는 부가세도 변경해야 함
                    VAT = CalculatePaymentListVAT(.selectedItem)
                    .selectedItem.ListSubItems(11).Text = Format(VAT, "#,##0")
                    UpdatePaymentListValue .selectedItem.Text, 12, VAT
                End If
                '합계
                Me.txtPaid.value = Format(CalculatePaymentListTotalCost, "#,##0")
                CalculatePayment
            End If
        End If
    End With
End Sub


Private Sub cboEstimatePayMethod_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswPaymentList
        If KeyCode = vbKeyReturn Then
            If headerIndex = 0 Then headerIndex = 10
            PaymentListUpdate headerIndex
            Me.cboEstimatePayMethod.Visible = False
            Me.frmPaymentEdit.Visible = False
            .SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If headerIndex = 0 Then headerIndex = 10
            PaymentListUpdate headerIndex
            headerIndex = headerIndex + 1
            SelectPaymentListColumn
            KeyCode = 0
        End If
    End With
End Sub

Private Sub txtPaymentEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long
    
    With Me.lswPaymentList
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
            '변경값을 DB와 화면에 반영
            PaymentListUpdate headerIndex

            '엔터키 - 값만 바꿔줌. 다음칸으로 이동하지 않음
            If KeyCode = vbKeyReturn Then
                Me.txtPaymentEdit.Visible = False
                Me.frmPaymentEdit.Visible = False

                Me.lswPaymentList.SetFocus
            ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyRight Then
                '탭키, 오른쪽 화살표키
                If headerIndex = 11 Then
                    Me.txtPaymentEdit.Visible = False
                    Me.frmPaymentEdit.Visible = False
                    Me.lswPaymentList.SetFocus
                Else
                    headerIndex = headerIndex + 1
                    SelectPaymentListColumn
                    KeyCode = 0
                End If
            ElseIf KeyCode = vbKeyUp Then
                '위쪽화살표키
                '리스트 맨 처음이 아니면 한칸위로 이동
                With Me.lswPaymentList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = 1 Then
                                Me.txtPaymentEdit.Visible = False
                                Me.frmPaymentEdit.Visible = False
                                Me.lswPaymentList.SetFocus
                            Else
                                .ListItems(i).Selected = False
                                .ListItems(i - 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectPaymentListColumn
                                KeyCode = 0
                                Exit For
                            End If
                        End If
                    Next
                End With
            ElseIf KeyCode = vbKeyDown Then
                '아래화살표키
                With Me.lswPaymentList
                    For i = 1 To .ListItems.count
                        If .ListItems(i).Selected = True Then
                            If i = .ListItems.count Then
                                '맨 마지막이면 마무리
                                Me.txtPaymentEdit.Visible = False
                                Me.frmPaymentEdit.Visible = False
                                Me.lswPaymentList.SetFocus
                                Exit For
                            Else
                                '리스트 맨 마지막이 아니면 한칸 아래로 이동
                                .ListItems(i).Selected = False
                                .ListItems(i + 1).Selected = True
                                Set beforeSelectedItem = .selectedItem
                                SelectPaymentListColumn
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
                    Me.txtPaymentEdit.Visible = False
                    Me.frmPaymentEdit.Visible = False
                    Me.lswPaymentList.SetFocus
                Else
                    headerIndex = headerIndex - 1
                    SelectPaymentListColumn
                    KeyCode = 0
                End If
            End If
        
        ElseIf KeyCode = vbKeyEscape Then
            'ESC키
            Me.txtPaymentEdit.Visible = False
            Me.frmPaymentEdit.Visible = False
        End If
    End With
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    '텍스트박스에 포커스를 두어야 AfterUpdate로 들어옴
    Me.txtEstimateDate.SetFocus
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtBidDate.SetFocus
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtInsuranceDate.SetFocus
    GetCalendarDate Me.txtInsuranceDate
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtAcceptedDate.SetFocus
    GetCalendarDate Me.txtAcceptedDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtDueDate.SetFocus
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.txtDeliveryDate.SetFocus
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxinvoiceDate
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "월")
End Sub


Private Sub txtManagementID_AfterUpdate()
    Dim li As ListItem
    
    '초기화 시에는 DB에 저장하는 것 방지
    If bInitialIzed = False Then Exit Sub
    
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
    
    '견적정보 DB 읽어오기
    db = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(db, Me.txtManagementID.value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbInformation, "작업 확인": Exit Sub
    
    '변경할 때마다 견적 DB에 저장
    UpdateEstimateOrderValue "관리번호", Me.txtManagementID.value
    
    '관리번호는 발주나간 건들도 DB에서 변경
    For Each li In lswOrderList.ListItems
    
    Next
    
    '관리번호 변경이 되는 경우 대비하여 바꿔줌
    orgManagementID = Me.txtManagementID.value
    
End Sub

Private Sub txtEstimateName_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtEstimateName.value = Trim(Me.txtEstimateName.value)
    
    UpdateEstimateOrderValue "견적명", Me.txtEstimateName.value
End Sub

Private Sub txtMemo_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtMemo.value = Trim(Me.txtMemo.value)
    
    UpdateEstimateOrderValue "메모", Me.txtMemo.value
End Sub

Private Sub txtAmount_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    If Me.txtAmount.value <> "" Then
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '수량 1,000자리 컴마 처리
            Me.txtAmount.value = Format(Me.txtAmount.value, "#,##0")
        End If
    End If
    
    '비용 필드 계산
    CalculateEstimatePrice
    
    'DB 적용
    UpdateEstimateOrderValue "수량", Me.txtAmount.value
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    If Me.txtUnitPrice.value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '단가 1,000자리 컴마 처리
            Me.txtUnitPrice.value = Format(Me.txtUnitPrice.value, "#,##0")
        End If
    End If
    
    '비용 필드 계산
    CalculateEstimatePrice
    
    'DB 적용
    '단가는 견적DB에만 저장
    UpdateEstimateValue "단가", Me.txtUnitPrice.value
End Sub

Private Sub txtEstimatePrice_Change()
    If bInitialIzed = False Then Exit Sub
    
    '견적금액은 견적DB만 변경함
     UpdateEstimateValue "금액", Me.txtEstimatePrice.value
End Sub

Private Sub txtBidPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     If Me.txtBidPrice.value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.value) Then
            Me.txtBidPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '입찰금액 1,000자리 컴마 처리
            Me.txtBidPrice.value = Format(Me.txtBidPrice.value, "#,##0")
        End If
    End If
    
    '비용 필드 계산
    CalculateBidMargin
    
    UpdateEstimateValue "입찰금액", Me.txtBidPrice.value
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     If Me.txtAcceptedPrice.value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtAcceptedPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            Me.txtAcceptedPrice.value = Format(Me.txtAcceptedPrice.value, "#,##0")
            
            CalculateAcceptedMargin
            CalculatePayment
            
            UpdateEstimateValue "수주금액", Me.txtAcceptedPrice.value
            UpdateOrderValue "금액", Me.txtAcceptedPrice.value
            If IsNumeric(Me.txtAmount.value) Then
                UpdateOrderValue "단가", CLng(Me.txtAcceptedPrice.value) / CLng(Me.txtAmount.value)
            End If
        End If
    End If
End Sub

Sub UpdateProductionTotalCost(fieldValue)
    Me.txtProductionTotalCost.value = fieldValue
End Sub

'Enable=False 인 텍스트박스는 AfterUpdate로 변경이 안되는 경우 있음. 이벤트를 Change로 바꿈
Private Sub txtProductionTotalCost_Change()
    If bInitialIzed = False Then Exit Sub
    
     '비용 필드 계산
    CalculateBidMargin
            
    'DB 반영
    UpdateEstimateOrderValue "실행가(예상)", Me.txtProductionTotalCost.value
End Sub

'Enable=False 인 텍스트박스는 AfterUpdate로 변경이 안되는 경우 있음. 이벤트를 Change로 바꿈
Private Sub txtExecutionCost_Change()
    If bInitialIzed = False Then Exit Sub
    
     CalculateAcceptedMargin
    
    UpdateEstimateOrderValue "실행가", Me.txtExecutionCost.value
End Sub

Private Sub txtPaid_Change()
    If bInitialIzed = False Then Exit Sub
    
    CalculatePayment
    
    UpdateEstimateOrderValue "입금액", Me.txtPaid.value
End Sub

Private Sub txtRemaining_Change()
    If bInitialIzed = False Then Exit Sub
    
    UpdateEstimateOrderValue "미입금액", Me.txtRemaining.value
End Sub

Private Sub txtVAT_Change()
    If bInitialIzed = False Then Exit Sub
    
    UpdateEstimateOrderValue "부가세", Me.txtVAT.value
End Sub

Private Sub chkVAT_AfterUpdate()
    If bInitialIzed = False Then Exit Sub

    If Me.chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        CalculatePayment
    End If
    
    UpdateEstimateOrderValue "부가세제외", Me.chkVAT.value
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtAcceptedDate.value = ConvertDateFormat(Me.txtAcceptedDate.value)
    
    UpdateEstimateOrderValue "수주", Me.txtAcceptedDate.value
End Sub

Private Sub cboCategory_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "분류1", Me.cboCategory.value
End Sub

Private Sub txtAcceptedMargin_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "수주차액", Me.txtAcceptedMargin.value
End Sub

Private Sub txtAcceptedMarginRate_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "마진율", Me.txtAcceptedMarginRate.value
End Sub

Private Sub txtBidMargin_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "차액(예상)", Me.txtBidMargin.value
End Sub

Private Sub txtBidMarginRate_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "마진율(예상)", Me.txtBidMarginRate.value
End Sub

Private Sub txtLinkedID_Change()
    If bInitialIzed = False Then Exit Sub
    
     UpdateEstimateOrderValue "연관번호", Me.txtLinkedID.value
End Sub
Private Sub cboUnit_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.cboUnit.value = Trim(Me.cboUnit.value)
    
    UpdateEstimateOrderValue "단위", Me.cboUnit.value
End Sub

Private Sub txtBidDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtBidDate.value = ConvertDateFormat(Me.txtBidDate.value)
    
    UpdateEstimateOrderValue "입찰", Me.txtBidDate.value
End Sub

Private Sub txtCustomer_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
    
    UpdateEstimateOrderValue "거래처", Me.txtCustomer.value
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtDeliveryDate.value = ConvertDateFormat(Me.txtDeliveryDate.value)
    
    UpdateEstimateOrderValue "납품", Me.txtDeliveryDate.value
End Sub

Private Sub txtDueDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtDueDate.value = ConvertDateFormat(Me.txtDueDate.value)
    
    UpdateEstimateOrderValue "납기", Me.txtDueDate.value
End Sub

Private Sub txtEstimateDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtEstimateDate.value = ConvertDateFormat(Me.txtEstimateDate.value)
    
    UpdateEstimateOrderValue "견적", Me.txtEstimateDate.value
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
     Me.txtInsuranceDate.value = ConvertDateFormat(Me.txtInsuranceDate.value)
    
    UpdateEstimateOrderValue "증권", Me.txtInsuranceDate.value
End Sub

Private Sub txtManager_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtManager.value = Trim(Me.txtManager.value)
    
    UpdateEstimateOrderValue "담당자", Me.txtManager.value
End Sub

Private Sub txtSize_AfterUpdate()
    If bInitialIzed = False Then Exit Sub
    
    Me.txtSize.value = Trim(Me.txtSize.value)
    
    UpdateEstimateOrderValue "규격", Me.txtSize.value
End Sub



Private Sub txtTaxinvoiceDate_AfterUpdate()
    Me.txtTaxinvoiceDate.value = Trim(Me.txtTaxinvoiceDate.value)
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = Trim(Me.txtPaymentDate.value)
End Sub

Private Sub txtExpectPaymentMonth_AfterUpdate()
    Dim pos As Long
    Dim M As String

    Me.txtExpectPaymentMonth.value = Trim(Me.txtExpectPaymentMonth.value)

    If Me.txtExpectPaymentMonth = "" Then Exit Sub
    
    pos = InStr(Me.txtExpectPaymentMonth, "월")
    If pos <> 0 Then
        M = Left(Me.txtExpectPaymentMonth, pos - 1)
        Me.txtExpectPaymentDate.value = DateSerial(Year(Date), M, 1)
        Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "월")
        Exit Sub
    End If
    
    If IsNumeric(Me.txtExpectPaymentMonth) Then
        Me.txtExpectPaymentDate.value = DateSerial(Year(Date), Me.txtExpectPaymentMonth, 1)
        Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "월")
        Exit Sub
    End If
    
    Me.txtExpectPaymentDate.value = Me.txtExpectPaymentMonth
    Me.txtExpectPaymentMonth.value = Format(Me.txtExpectPaymentDate.value, "mm" & "월")
     
End Sub

Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = Trim(Me.txtSpecificationDate.value)
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

Private Sub txtPaymentEdit_AfterUpdate()
    '탭키나 엔터키가 아닌 마우스를 클릭해서 벗어나는 경우: currentEditText를 사용해야 함
    If headerIndex >= 4 And headerIndex < Me.lswPaymentList.ColumnHeaders.count Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.txtPaymentEdit.value <> currentEditText Then
                PaymentListUpdate headerIndex
                headerIndex = 0
                currentEditText = ""
            End If
        End If
    End If
End Sub

Private Sub cboEstimatePayMethod_AfterUpdate()
    If headerIndex = 10 Then
        If Not beforeSelectedItem Is Nothing Then
            If Me.cboEstimatePayMethod.value <> currentCboText Then
                PaymentListUpdate headerIndex
                headerIndex = 0
                currentCboText = ""
            End If
        End If
    End If
End Sub

Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub


Private Sub UserForm_Terminate()
    bInitialIzed = False
End Sub
