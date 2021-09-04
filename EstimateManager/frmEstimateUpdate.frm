VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "견적 수정"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19200
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




Private Sub cboUnit_AfterUpdate()
    Me.cboUnit.Value = Trim(Me.cboUnit.Value)
End Sub


Private Sub txtBidDate_AfterUpdate()
    Me.txtBidDate.Value = Trim(Me.txtBidDate.Value)
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.Value = Trim(Me.txtCustomer.Value)
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    Me.txtDeliveryDate.Value = Trim(Me.txtDeliveryDate.Value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.Value = Trim(Me.txtDueDate.Value)
End Sub


Private Sub txtEstimateDate_AfterUpdate()
    Me.txtEstimateDate.Value = Trim(Me.txtEstimateDate.Value)
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    Me.txtInsuranceDate.Value = Trim(Me.txtInsuranceDate.Value)
End Sub

Private Sub txtManager_AfterUpdate()
    Me.txtManager.Value = Trim(Me.txtManager.Value)
End Sub


Private Sub txtSize_AfterUpdate()
    Me.txtSize.Value = Trim(Me.txtSize.Value)
End Sub


Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.Value = Trim(Me.txtSpecificationDate.Value)
End Sub

Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim db As Variant
    Dim contr As Control
    
    If clickEstimateId <> "" Then              '발주관리에서 더블클릭한 경우
        currentEstimateId = CLng(clickEstimateId)
        clickEstimateId = ""
    Else
        '선택한 행 번호
        cRow = Selection.row
    
        '데이터가 있는 행이 아닐 경우는 중지
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then
            MsgBox "수정할 견적 행을 먼저 선택한 후 견적수정 버튼을 클릭하세요."
            End
        End If
        
        currentEstimateId = shtEstimateAdmin.Cells(cRow, 2)
    End If
    
     '텍스트박스 라벨 컨트롤 색상 조정
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            'contr.top = contr.top + 2
            If contr.Name Like "lbl2*" Then
                
            Else
                contr.BackColor = RGB(242, 242, 242)
            End If
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

    Me.txtID.Value = estimate(1)    'ID
    Me.txtEstimateName.Value = estimate(6)  '견적명
    Me.txtManagementID.Value = estimate(2)    '관리번호
    Me.txtLinkedID.Value = estimate(3)  '자재번호
    
    Me.txtCustomer = estimate(4)   '거래처
    Me.txtManager = estimate(5)   '담당자
    
    Me.txtSize.Value = estimate(7)  '규격
    Me.txtAmount.Value = Format(estimate(8), "#,##0")   '수량
    InitializeCboUnit
    Me.cboUnit.Value = Trim(estimate(9))  '단위, ID가 없으므로 직접 value 넣으면 선택됨
    Me.txtUnitPrice.Value = Format(estimate(10), "#,##0")     '견적단가
    Me.txtEstimatePrice.Value = Format(estimate(11), "#,##0")     '견적금액
    
    Me.txtEstimateDate.Value = estimate(12)    '견적일자
    Me.txtBidDate.Value = estimate(13)    '입찰일자
    Me.txtAcceptedDate.Value = estimate(14)    '수주일자
    Me.txtDeliveryDate.Value = estimate(15)    '납품일자
    Me.txtInsuranceDate.Value = estimate(16)    '증권일자
    
    Me.txtProductionTotalCost.Value = Format(estimate(17), "#,##0")   '실행가
    Me.txtBidPrice.Value = Format(estimate(18), "#,##0")    '입찰가
    Me.txtBidMargin.Value = Format(estimate(19), "#,##0")    '차액
    Me.txtBidMarginRate.Value = Format(estimate(20), "0.0%")    '마진율
    Me.txtAcceptedPrice.Value = Format(estimate(21), "#,##0")    '수주금액
    Me.txtAcceptedMargin.Value = Format(estimate(22), "#,##0")   '수주차액
    
    Me.txtInsertDate.Value = estimate(23)    '등록일자
    Me.txtUpdateDate.Value = estimate(24)    '수정일자
    
    InitializeCboCategory
    Me.cboCategory.Value = Trim(estimate(25))   '분류
    '26은 납기일
    Me.txtSpecificationDate.Value = estimate(27)    '거래명세서
    Me.txtTaxInvoiceDate.Value = estimate(28)    '세금계산서
    Me.txtPaymentDate.Value = estimate(29)    '결제일자
    Me.txtExpectPaymentDate.Value = estimate(30)  '예상결제일
    Me.txtExpectPaymentMonth.Value = Format(estimate(30), "mm" & "월")  '예상결제월
    Me.txtVAT.Value = Format(estimate(31), "#,##0")    '부가세
    Me.txtMemo.Value = estimate(32)
    Me.chkVAT.Value = estimate(33)      '부가세 제외 여부
    
    Me.txtPaid.Value = Format(estimate(34), "#,##0")      '입금액
    Me.txtRemaining.Value = Format(estimate(35), "#,##0")      '미입금액
    Me.chkDividePay.Value = estimate(36)      '분할결제 여부
    
    '변경 전 관리번호
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '발주 현황
    
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
    db = Filtered_DB(db, Me.cboCustomer.Value, 1, True)
    
    '기존 콤보박스 내용지우기
    Me.cboManager.Clear
    
    '담당자가 있으면 콤보박스에 추가함
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


Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '견적ID에 해당하는 발주 정보를 읽어옴
    db = Get_DB(shtOrder)
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.Value, 28, True)
    End If
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, "<>" & "수주", 4)
    End If
    
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
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "품명", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "분류", 34
        .ColumnHeaders.Add , , "거래처", 70
        .ColumnHeaders.Add , , "재질", 62
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
        .ColumnHeaders.Add , , "결제일", 59, lvwColumnCenter
        
        .ColumnHeaders(1).Position = 6
    
        .ListItems.Clear
        totalCost = 0
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 7))   '품명
                li.ListSubItems.Add , , db(i, 1)        'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_견적
                li.ListSubItems.Add , , db(i, 5)        '관리번호
                li.ListSubItems.Add , , db(i, 4)        '분류
                li.ListSubItems.Add , , db(i, 6)        '거래처
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
                li.ListSubItems.Add , , db(i, 22)       '결제일
                li.Selected = False
                
                If IsNumeric(db(i, 13)) Then
                    '비용 합계 구함
                    totalCost = totalCost + CLng(db(i, 13))
                End If
            Next
        End If
        
        If totalCost <> 0 Then
            Me.txtExecutionCost.Value = Format(totalCost, "#,##0")
        End If
    End With
End Sub

Sub UpdateEstimate()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckEstimateUpdateValidation = False Then
        Exit Sub
    End If

    '견적정보 DB 읽어오기
    db = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    '데이터 업데이트
    Update_Record shtEstimate, Me.txtID.Value, _
        Me.txtManagementID.Value, Me.txtLinkedID.Value, _
        Me.txtCustomer.Value, Me.txtManager.Value, _
        Me.txtEstimateName.Value, Me.txtSize.Value, _
        Me.txtAmount.Value, Me.cboUnit.Value, _
        Me.txtUnitPrice.Value, Me.txtEstimatePrice.Value, _
        Me.txtEstimateDate.Value, Me.txtBidDate.Value, _
        Me.txtAcceptedDate.Value, Me.txtDeliveryDate.Value, _
        Me.txtInsuranceDate.Value, Me.txtProductionTotalCost.Value, _
        Me.txtBidPrice.Value, Me.txtBidMargin.Value, _
        Me.txtBidMarginRate.Value, Me.txtAcceptedPrice.Value, _
        Me.txtAcceptedMargin.Value, _
        Me.txtInsertDate.Value, Date, _
        Me.cboCategory.Value, Me.txtDueDate.Value, _
        Me.txtSpecificationDate.Value, Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, Me.txtExpectPaymentDate.Value, _
        Me.txtVAT.Value, Me.txtMemo.Value, Me.chkVAT.Value, _
        Me.txtPaid.Value, Me.txtRemaining.Value, Me.chkDividePay
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.Range("H" & selectionRow).Select
End Sub

Function CheckEstimateUpdateValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '견적명이 입력되었는지 체크
    If Me.txtEstimateName.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "견적명을 입력하세요."
    End If
    
    '관리번호가 입력되었는지 체크
    If Me.txtManagementID.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "관리번호를 입력하세요."
    End If
    
    If bCorrect = False Then
        Me.lblErrorMessage.Visible = True
    Else
        Me.lblErrorMessage.Visible = False
    End If
    
    CheckEstimateUpdateValidation = bCorrect
End Function

Sub CalculateEstimateUpdateCost()

    '견적금액 계산
    '수량값이 공백이면 견적금액은 견적단가
    If Me.txtUnitPrice <> "" Then
        If Me.txtAmount.Value = "" Then
            Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
        Else
            Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")

    '예상차액과 예상마진율 계산
    If Me.txtBidPrice.Value <> "" And Me.txtProductionTotalCost.Value <> "" Then
        '예상차액 = 입찰가 - 예상실행가
        Me.txtBidMargin.Value = Format(CLng(Me.txtBidPrice.Value) - CLng(Me.txtProductionTotalCost.Value), "#,##0")
        '예상마진율 = 예상차액 / 입찰가
        If Me.txtBidPrice.Value <> "0" Then
            Me.txtBidMarginRate.Value = Format(CLng(Me.txtBidMargin.Value) / CLng(Me.txtBidPrice.Value), "0.0%")
        End If
    Else
        Me.txtBidMargin.Value = 0
    End If

    '수주차액, 마진율 계산
    If Me.txtAcceptedPrice.Value <> "" And Me.txtExecutionCost.Value <> "" Then
        '수주차액 = 수주금액 - 실행가
        Me.txtAcceptedMargin.Value = Format(CLng(Me.txtAcceptedPrice.Value) - CLng(Me.txtExecutionCost.Value), "#,##0")
        '마진율 = 수주차액 / 수주금액
        If Me.txtAcceptedPrice.Value <> "0" Then
            Me.txtAcceptedMarginRate.Value = Format(CLng(Me.txtAcceptedMargin.Value) / CLng(Me.txtAcceptedPrice.Value), "0.0%")
        End If
    Else
        Me.txtAcceptedMargin.Value = ""
        Me.txtAcceptedMarginRate.Value = ""
    End If

    '부가세 계산
    '세금계산서 일자가 없는 경우, 부가세 제외인 경우 부가세는 0
    If Me.txtTaxInvoiceDate.Value = "" Or chkVAT.Value = True Then
        Me.txtVAT.Value = 0
    Else
        '부가세는 수주금액의 10%
        If Me.txtAcceptedPrice.Value <> "" And Me.txtAcceptedPrice.Value <> 0 Then
            Me.txtVAT.Value = CLng(Me.txtAcceptedPrice.Value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.Value, "#,##0")
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


Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2, True)
    
    'DB에 값이 있을 경우
    totalCost = 0
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 11)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(db(i, 11))
            End If
        Next
    End If
        
    GetProductionTotalCost = totalCost
End Function


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
    UpdateEstimate
End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnOrderListDelete_Click()
    Dim li As ListItem
    Dim count As Long
    Dim YN As VbMsgBoxResult
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "삭제할 발주를 선택하세요.": Exit Sub
    
    YN = MsgBox("선택한 " & count & "개 발주를 삭제합니다.", vbYesNo)
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Then
            '발주 테이블에서 삭제
            Delete_Record shtOrder, li.SubItems(1)
        End If
    Next
    
    If count > 0 Then
        InitializeLswOrderList
    End If
End Sub

Private Sub btnProduction_Click()
    If isFormLoaded("frmProduction") Then
        Unload frmProduction
    End If
    frmProduction.Show (False)
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub


Private Sub btnPayHistoryInsert_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnEstimateUpdate.SetFocus
    End If
End Sub

Private Sub btnProduction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub

Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.txtAcceptedDate.SetFocus
    End If
End Sub

Private Sub lswOrderCustomerAutoComplete_DblClick()
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
    With Me.lswOrderCustomerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtProductionCustomer.Value = .SelectedItem.Text
            .Visible = False
            Me.txtProductionItem.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
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

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtDeliveryDate
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
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "월")
End Sub

Private Sub cboCustomer_Change()
    InitializeCboManager
End Sub

Private Sub txtEstimateDate_Change()
    
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.txtManagementID.Value = Trim(Me.txtManagementID.Value)
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.txtEstimateName.Value = Trim(Me.txtEstimateName.Value)
End Sub

Private Sub txtAmount_AfterUpdate()

    If Me.txtAmount.Value <> "" Then
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '수량 1,000자리 컴마 처리
            Me.txtAmount.Value = Format(Me.txtAmount.Value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.Value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '견적단가 1,000자리 컴마 처리
            Me.txtUnitPrice.Value = Format(Me.txtUnitPrice.Value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtBidPrice_AfterUpdate()
    
    If Me.txtBidPrice.Value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtBidPrice.Value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '입찰금액 1,000자리 컴마 처리
            Me.txtBidPrice.Value = Format(Me.txtBidPrice.Value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If Me.txtAcceptedPrice.Value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.Value) Then
            Me.txtAcceptedPrice.Value = ""
            MsgBox "숫자를 입력하세요."
        Else
            Me.txtAcceptedPrice.Value = Format(Me.txtAcceptedPrice.Value, "#,##0")
            
            CalculateEstimateUpdateCost
        End If
    End If
End Sub

Private Sub txtProductionTotalCost_AfterUpdate()
    
    If Me.txtProductionTotalCost.Value <> "" Then
        If Not IsNumeric(Me.txtProductionTotalCost.Value) Then
            Me.txtProductionTotalCost.Value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '예상실행가 1,000자리 컴마 처리
            Me.txtProductionTotalCost.Value = Format(Me.txtProductionTotalCost.Value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtExecutionCost_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    Me.txtAcceptedDate.Value = Trim(Me.txtAcceptedDate.Value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
    Me.txtTaxInvoiceDate.Value = Trim(Me.txtTaxInvoiceDate.Value)
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.Value = Trim(Me.txtPaymentDate.Value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtExpectPaymentMonth_AfterUpdate()
    Dim pos As Long
    Dim M As String

    Me.txtExpectPaymentMonth.Value = Trim(Me.txtExpectPaymentMonth.Value)

    If Me.txtExpectPaymentMonth = "" Then Exit Sub
    
    pos = InStr(Me.txtExpectPaymentMonth, "월")
    If pos <> 0 Then
        M = Left(Me.txtExpectPaymentMonth, pos - 1)
        Me.txtExpectPaymentDate.Value = DateSerial(Year(Date), M, 1)
        Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "월")
        Exit Sub
    End If
    
    If IsNumeric(Me.txtExpectPaymentMonth) Then
        Me.txtExpectPaymentDate.Value = DateSerial(Year(Date), Me.txtExpectPaymentMonth, 1)
        Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "월")
        Exit Sub
    End If
    
    Me.txtExpectPaymentDate.Value = Me.txtExpectPaymentMonth
    Me.txtExpectPaymentMonth.Value = Format(Me.txtExpectPaymentDate.Value, "mm" & "월")
     
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub

