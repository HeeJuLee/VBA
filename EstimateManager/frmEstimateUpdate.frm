VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "견적 정보 수정"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17730
   OleObjectBlob   =   "frmEstimateUpdate.frx":0000
   StartUpPosition =   1  '소유자 가운데
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
    '콤보박스에서 거래처를 변경하면 해당 거래처의 담당자로 담당자 콤보박스를 세팅
    InitializeCboManager
End Sub

'수주일자 입력 박스
Private Sub txtAcceptedDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'증권보험 일자 입력박스
Private Sub txtInsuranceDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'거래명세서 일자 입력박스
Private Sub txtSpecificationDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'세금계산서 일자 입력박스
Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateEstimateUpdateCost
End Sub

'결제일자 입력박스
Private Sub txtPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

'예상결제일자 입력박스
Private Sub txtExpectPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtEstimateDate_Change()
    '오류 메시지 숨김
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateID_AfterUpdate()
    '오류 메시지 숨김
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    '오류 메시지 숨김
    Me.lblErrorMessage.Visible = False
End Sub

'수량 입력
Private Sub txtAmount_AfterUpdate()
    '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    '수량값이 공백이면 중지
    If Me.txtAmount.Value <> "" Then
        '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
End Sub

'견적단가 입력
Private Sub txtUnitPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '견적단가 1,000자리 컴마 처리
    Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
End Sub

'예상 실행 시뮬레이션 비용 입력
Private Sub txtProductionCost_AfterUpdate()
    '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtProductionCost.Value = "" Then
        Exit Sub
    End If
    
    '비용 입력값이 숫자가 아닐 경우 오류메시지 출력
    If Not IsNumeric(Me.txtProductionCost.Value) Then
        Me.txtProductionCost.Value = ""
        Me.lblErrorMessage.Caption = "숫자를 입력하세요."
        Me.lblErrorMessage.Visible = True
        Exit Sub
    End If
    
    '합계 금액 1,000자리 컴마 처리
    Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
End Sub

'예상실행가 입력
Private Sub txtProductionTotalCost_AfterUpdate()
     '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtProductionTotalCost.Value <> "" Then
        '예상실행가 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtProductionTotalCost.Value) Then
            Me.txtProductionTotalCost.Value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '예상 실행 금액 1,000자리 컴마 처리
    Me.txtProductionTotalCost.Text = Format(Me.txtProductionTotalCost.Value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
End Sub

'입찰금액 입력
Private Sub txtBidPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtBidPrice.Value <> "" Then
        '입찰금액이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtBidPrice.Value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If

    '입찰금액 1,000자리 컴마 처리
    Me.txtBidPrice.Text = Format(Me.txtBidPrice.Value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
    
End Sub

'수주일자 캘린더 선택
Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

'입찰일자 캘린더 선택
Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

'납품일자 캘린더 선택
Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

'견적일자 캘린더 선택
Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

'증권보험 캘린더 선택
Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

'결제일자 캘린더 선택
Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
    CalculateEstimateUpdateCost
End Sub

'거래명세서 캘린더 선택
Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

'세금계산서 캘린더 선택
Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateEstimateUpdateCost
End Sub

'예상결제일자 캘린더 선택
Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    CalculateEstimateUpdateCost
End Sub


Private Sub lstProductionList_Click()
    Dim arr As Variant

    arr = Get_ListItm(Me.lstProductionList)
    
    Me.txtProductionID.Value = arr(0)
    Me.txtProductionItem.Value = arr(2)
    Me.txtProductionCost.Value = arr(3)
    Me.txtProductionCost.Text = Format(arr(3), "#,##0")
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
    
    '선택한 행 번호
    cRow = Selection.row

    '데이터가 있는 행이 아닐 경우는 중지
    If cRow < 8 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then End
    
    '견적/담당자/거래처 데이터 읽어오기
    estimate = Get_Record_Array(shtEstimate, shtEstimateAdmin.Cells(cRow, 2))
    manager = Get_Record_Array(shtManager, estimate(2))
    customer = Get_Record_Array(shtCustomer, manager(2))

    Me.txtID.Value = estimate(1)    'ID
    Me.txtManagerID.Value = estimate(2) 'ID_담당자
    Me.txtEstimateName.Value = estimate(5)  '견적명
    Me.txtEstimateID.Value = estimate(3)    '관리번호
    Me.txtLinkedID.Value = estimate(4)  '자재번호
    
    InitializeCboCustomer
    Select_CboItm Me.cboCustomer, customer(1), 1    '거래처
    InitializeCboManager
    Select_CboItm Me.cboManager, manager(1), 1  '담당자
    
    Me.txtSize.Value = estimate(6)  '규격
    
    InitializeCboUnit
    Me.cboUnit.Value = estimate(8)  '단위, ID가 없으므로 직접 value 넣으면 선택됨
    
    Me.txtAmount.Value = Format(estimate(7), "#,##0")   '수량
    Me.txtUnitPrice.Value = Format(estimate(9), "#,##0")     '견적단가
    Me.txtEstimatePrice.Value = Format(estimate(10), "#,##0")     '견적금액
    
    Me.txtEstimateDate.Value = estimate(11)    '견적일자
    Me.txtBidDate.Value = estimate(12)    '입찰일자
    Me.txtAcceptedDate.Value = estimate(13)    '수주일자
    Me.txtDeliveryDate.Value = estimate(14)    '납품일자
    Me.txtInsuranceDate.Value = estimate(15)    '증권일자
    
    InitializeLstProduction    '예상실행 입력목록
    Me.txtProductionTotalCost.Value = Format(estimate(16), "#,##0")    '예상실행가
    
    Me.txtBidPrice.Value = Format(estimate(17), "#,##0")    '입찰가
    Me.txtBidMargin.Value = Format(estimate(18), "#,##0")    '차액
    Me.txtBidMarginRate.Value = Format(estimate(19), "0.0%")    '마진율
    Me.txtAcceptedPrice.Value = Format(estimate(20), "#,##0")    '수주금액
    Me.txtAcceptedMargin.Value = Format(estimate(21), "#,##0")   '수주차액
    
    Me.txtSpecificationDate.Value = estimate(22)    '거래명세서
    Me.txtTaxInvoiceDate.Value = estimate(23)    '세금계산서
    Me.txtPaymentDate.Value = estimate(24)    '결제일자
    Me.txtExpectPaymentDate.Value = estimate(25)    '예상결제일자
    Me.txtVAT.Value = Format(estimate(26), "#,##0")    '부가세
    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '입금예상액
    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '입금액
    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '미입금액
    
    Me.txtInsertDate.Value = estimate(30)    '등록일자
    Me.txtUpdateDate.Value = estimate(31)    '수정일자
    

    '변경 전 관리번호
    orgEstimateID = Me.txtEstimateID
    
    
End Sub


Sub UpdateEstimate()
    Dim DB As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckEstimateUpdateValidation = False Then
        Exit Sub
    End If

    '견적정보 DB 읽어오기
    DB = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(DB, Me.txtEstimateID.Value, 3, orgEstimateID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    '견적금액 계산 = 견적단가 * 수량
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value * Me.txtAmount.Value
    End If
    
    '데이터 업데이트
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
    
    '담당자 DB를 읽어와서
    DB = Get_DB(shtManager)
    '거래처ID로 필터링
    DB = Filtered_DB(DB, Me.cboCustomer.Value, 2)
    
    '기존 콤보박스 내용지우기
    Me.cboManager.Clear
    
    '담당자가 있으면 콤보박스에 추가함
    If Not IsEmpty(DB) Then
        'Filtered_DB 통과하면서 ID가 문자열로 바뀜 -> 이걸 숫자로 변환
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
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    DB = Get_DB(shtProduction)
    DB = Filtered_DB(DB, Me.txtID.Value, 2)
    
    'DB에 값이 있을 경우
    If Not IsEmpty(DB) Then
        For i = 1 To UBound(DB)
            If IsNumeric(DB(i, 4)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(DB(i, 4))
                '숫자 포맷 1,000자리 처리
                DB(i, 4) = Format(DB(i, 4), "#,##0")
            End If
        Next
        
        Me.txtProductionSum = Format(totalCost, "#,##0")
        
        Update_List Me.lstProductionList, DB, "0pt;0pt;60pt;50pt;100pt;"
    End If
    
End Sub

Sub InitalizeProductionInput()
    Me.txtProductionID.Value = ""
    Me.txtProductionItem.Value = ""
    Me.txtProductionCost.Value = ""
    Me.txtProductionMemo.Value = ""
End Sub

Sub InsertProjection()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "항목을 입력하세요.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "비용을 입력하세요.": Exit Sub
    
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

    If Me.txtProductionItem.Value = "" Then MsgBox "항목을 입력하세요.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "비용을 입력하세요.": Exit Sub

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

Function CheckEstimateUpdateValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '견적명이 입력되었는지 체크
    If Me.txtEstimateName.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "견적명을 입력하세요."
    End If
    
    '관리번호가 입력되었는지 체크
    If Me.txtEstimateID.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "관리번호를 입력하세요."
    End If
    
    '수량이 입력되었는지 체크
    If Me.txtAmount.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "수량을 입력하세요."
    End If
    
    '견적단가가 입력되었는지 체크
    If Me.txtUnitPrice.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "견적단가를 입력하세요."
    End If
    
    '견적일자가 입력되었는지 체크
    If Me.txtEstimateDate.Value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "견적일자를 입력하세요."
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
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")

    '차액과 마진율 계산
    If Me.txtBidPrice.Value <> "" And Me.txtProductionTotalCost <> "" Then
        '차액 = 입찰금액 - 예상실행금액
        Me.txtBidMargin.Value = CLng(Me.txtBidPrice.Value) - CLng(Me.txtProductionTotalCost.Value)
        Me.txtBidMargin.Text = Format(Me.txtBidMargin.Value, "#,##0")
        '마진율 = 차액 / 입찰금액
        Me.txtBidMarginRate.Value = CLng(Me.txtBidMargin.Value) / CLng(Me.txtBidPrice.Value)
        Me.txtBidMarginRate.Text = Format(Me.txtBidMarginRate.Value, "0.0%")
    Else
        Me.txtBidMargin.Value = 0
    End If

    '수주금액 계산
    If Me.txtAcceptedDate.Value = "" Then
        '수주일자가 없는 경우
        Me.txtAcceptedPrice.Value = 0
        Me.txtAcceptedMargin.Value = 0
    Else
        '수주일자가 있는 경우
        '수주금액은 입찰금액으로 세팅
        If IsNumeric(Me.txtBidPrice.Value) Then
            Me.txtAcceptedPrice.Value = CLng(Me.txtBidPrice.Value)
        Else
            Me.txtAcceptedPrice.Value = 0
        End If
        Me.txtAcceptedPrice.Text = Format(Me.txtAcceptedPrice.Value, "#,##0")
        
        '수주차액은 차액으로 세팅
        If IsNumeric(Me.txtBidMargin.Value) Then
            Me.txtAcceptedMargin.Value = CLng(Me.txtBidMargin.Value)
        Else
            Me.txtAcceptedMargin.Value = 0
        End If
        Me.txtAcceptedMargin.Text = Format(Me.txtAcceptedMargin.Value, "#,##0")
    End If

    '부가세 계산
    '세금계산서 일자가 있는 경우만
    If Me.txtTaxInvoiceDate.Value = "" Then
        Me.txtVAT.Value = 0
    Else
        '부가세는 수주금액의 10%
        If Me.txtAcceptedPrice.Value <> "" And Me.txtAcceptedPrice.Value <> 0 Then
            Me.txtVAT.Value = CLng(Me.txtAcceptedPrice.Value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.Value, "#,##0")
        End If
    End If

    '입금예상액 계산
    If Me.txtTaxInvoiceDate.Value = "" Then
        '세금계산서 일자가 없는 경우는 수주금액
        Me.txtExpectPay.Value = Me.txtAcceptedPrice
    Else
        '세금계산서 일자가 있는 경우는 수주금액+부가세
        If Me.txtAcceptedPrice.Value <> "" Then
            Me.txtExpectPay.Value = CLng(Me.txtAcceptedPrice.Value) + CLng(Me.txtVAT.Value)
        End If
    End If
    Me.txtExpectPay.Text = Format(Me.txtExpectPay.Value, "#,##0")

    '입금액 계산
    If Me.txtPaymentDate.Value = "" Then
        Me.txtPaid.Value = 0
    Else
        Me.txtPaid.Value = Me.txtExpectPay.Value
        Me.txtPaid.Text = Format(Me.txtPaid.Value, "#,##0")
    End If
    
    '미입금액 계산
    Me.txtUnpaid.Value = CLng(Me.txtExpectPay.Value) - CLng(Me.txtPaid.Value)
    Me.txtUnpaid.Text = Format(Me.txtUnpaid.Value, "#,##0")
    
End Sub

'=============================================
'리스트박스 스크롤
'Private Sub lstProductionList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstProductionList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstProductionList
'End Sub


