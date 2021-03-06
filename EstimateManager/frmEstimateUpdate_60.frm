VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate_60 
   Caption         =   "견적 수정"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18720
   OleObjectBlob   =   "frmEstimateUpdate_60.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEstimateUpdate_60"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim orgEstimateID As Variant


Private Sub btnEstimateClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
    
    '견적관리 화면 새로고침
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
End Sub

Private Sub btnEstimateUpdate_Click()
    UpdateEstimate
End Sub

Private Sub btnProductionClear_Change()

End Sub

Private Sub btnProductionClear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearProductionInput
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

'부가세 체크박스
Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
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

'예상실행가 입력
'Private Sub txtProductionTotalCost_AfterUpdate()
'     '오류메시지 숨김
'    Me.lblErrorMessage.Visible = False
'
'    If Me.txtProductionTotalCost.Value <> "" Then
'        '예상실행가 숫자가 아닐 경우 오류메시지 출력
'        If Not IsNumeric(Me.txtProductionTotalCost.Value) Then
'            Me.txtProductionTotalCost.Value = ""
'            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
'            Me.lblErrorMessage.Visible = True
'        End If
'    End If
'
'    '예상 실행 금액 1,000자리 컴마 처리
'    Me.txtProductionTotalCost.Text = Format(Me.txtProductionTotalCost.Value, "#,##0")
'
'    '비용 필드 계산
'    CalculateEstimateUpdateCost
'End Sub

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

'예상실행항목 수량 입력
Private Sub txtProductionAmount_AfterUpdate()
    '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtProductionAmount.Value = "" Then
        Exit Sub
    End If
    
    If IsNumeric(Me.txtProductionAmount.Value) Then
        Me.txtProductionAmount.Text = Format(Me.txtProductionAmount.Value, "#,##0")
        
        '금액 = 수량 * 단가
        If IsNumeric(Me.txtProductionUnitPrice.Value) Then
            Me.txtProductionCost.Value = CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value)
            Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
        End If
    End If
End Sub

'예상실행항목 단가 입력
Private Sub txtProductionUnitPrice_AfterUpdate()
    '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
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
                '금액 = 수량 * 단가
                Me.txtProductionCost.Value = CLng(Me.txtProductionAmount.Value) * CLng(Me.txtProductionUnitPrice.Value)
                Me.txtProductionCost.Text = Format(Me.txtProductionCost.Value, "#,##0")
            End If
        End If
    End If
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
    Me.txtProductionID.Value = arr(0)                       'ID
    Me.txtProductionCustomer = arr(3)               '거래처
    Me.txtProductionItem.Value = arr(4)                     '품명
    Me.txtProductionMaterial.Value = arr(5)           '재질
    Me.txtProductionSize.Value = arr(6)                '규격
    Me.txtProductionAmount.Value = arr(7)           '수량
    Me.cboProductionUnit.Value = arr(8)               '단위
    Me.txtProductionUnitPrice.Value = arr(9)        '단가
    Me.txtProductionUnitPrice.Text = Format(arr(9), "#,##0")
    Me.txtProductionCost.Value = arr(10)         '금액
    Me.txtProductionCost.Text = Format(arr(10), "#,##0")
    Me.txtProductionMemo = arr(11)       '메모
    
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim manager As Variant
    Dim customer As Variant
    Dim db As Variant
    
    '선택한 행 번호
    cRow = Selection.row

    '데이터가 있는 행이 아닐 경우는 중지
    If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).Value = "" Then End
    
    '견적/담당자/거래처 데이터 읽어오기
    estimate = Get_Record_Array(shtEstimate, shtEstimateAdmin.Cells(cRow, 2))

    Me.txtID.Value = estimate(1)    'ID
    Me.txtEstimateName.Value = estimate(6)  '견적명
    Me.txtEstimateID.Value = estimate(2)    '관리번호
    Me.txtLinkedID.Value = estimate(3)  '자재번호
    
    InitializeCboCustomer
    Select_CboItm Me.cboCustomer, Trim(estimate(4)), 1    '거래처
    InitializeCboManager
    Select_CboItm Me.cboManager, Trim(estimate(5)), 1  '담당자
    
    Me.txtSize.Value = estimate(7)  '규격
    
    InitializeCboUnit
    Me.cboUnit.Value = Trim(estimate(9))  '단위, ID가 없으므로 직접 value 넣으면 선택됨
    
    Me.txtAmount.Value = Format(estimate(8), "#,##0")   '수량
    Me.txtUnitPrice.Value = Format(estimate(10), "#,##0")     '견적단가
    Me.txtEstimatePrice.Value = Format(estimate(11), "#,##0")     '견적금액
    
    Me.txtEstimateDate.Value = estimate(12)    '견적일자
    Me.txtBidDate.Value = estimate(13)    '입찰일자
    Me.txtAcceptedDate.Value = estimate(14)    '수주일자
    Me.txtDeliveryDate.Value = estimate(15)    '납품일자
    Me.txtInsuranceDate.Value = estimate(16)    '증권일자
    
    InitializeLstProduction    '예상실행항목 목록
    InitializeCboProductonUnit  '예상실행항목 단위
    
    Me.txtBidPrice.Value = Format(estimate(18), "#,##0")    '입찰가
    Me.txtBidMargin.Value = Format(estimate(19), "#,##0")    '차액
    Me.txtBidMarginRate.Value = Format(estimate(20), "0.0%")    '마진율
    Me.txtAcceptedPrice.Value = Format(estimate(21), "#,##0")    '수주금액
    Me.txtAcceptedMargin.Value = Format(estimate(22), "#,##0")   '수주차액
    
    InitializeCboCategory
    Me.cboCategory.Value = Trim(estimate(25))   '분류
    Me.txtSpecificationDate.Value = estimate(26)    '거래명세서
    Me.txtTaxInvoiceDate.Value = estimate(27)    '세금계산서
    Me.txtPaymentDate.Value = estimate(28)    '결제일자
    Me.txtExpectPaymentDate.Value = estimate(29)    '예상결제일자
    Me.txtVAT.Value = Format(estimate(30), "#,##0")    '부가세
    Me.chkVAT.Value = estimate(31)
    
'    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '입금예상액
'    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '입금액
'    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '미입금액
    
    Me.txtInsertDate.Value = estimate(23)    '등록일자
    Me.txtUpdateDate.Value = estimate(24)    '수정일자
    
    
    '변경 전 관리번호
    orgEstimateID = Me.txtEstimateID
    
    
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
    blnUnique = IsUnique(db, Me.txtEstimateID.Value, 3, orgEstimateID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    '견적금액 계산 = 견적단가 * 수량
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
    Else
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value * Me.txtAmount.Value
    End If
    
    '데이터 업데이트
    Update_Record shtEstimate, Me.txtID.Value, _
        Me.txtEstimateID.Value, Me.txtLinkedID.Value, _
        Me.cboCustomer.Value, Me.cboManager.Value, _
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
        Me.cboCategory.Value, Me.txtSpecificationDate.Value, _
        Me.txtTaxInvoiceDate.Value, Me.txtPaymentDate.Value, _
        Me.txtExpectPaymentDate.Value, Me.txtVAT.Value, Me.chkVAT.Value

    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    
End Sub

Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtCustomer, True)

    Update_Cbo Me.cboCustomer, db
End Sub

Sub InitializeCboCustomer2()
    Dim db As Variant
    db = Get_DB(shtCustomer, True)

    Update_Cbo Me.cboCustomer2, db
End Sub

Sub InitializeCboManager()
    Dim db As Variant
    Dim i As Long
    
    '담당자 DB를 읽어와서
    db = Get_DB(shtManager, True)
    '거래처명으로 필터링
    db = Filtered_DB(db, Me.cboCustomer.Value, 1)
    
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

Sub InitializeLstProduction()
    Dim db As Variant
    Dim i, totalCost As Long
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2)
    
    'DB에 값이 있을 경우
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 10)) Then
                db(i, 10) = Format(db(i, 10), "#,##0")
            End If
            If IsNumeric(db(i, 11)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(db(i, 11))
                '숫자 포맷 1,000자리 처리
                db(i, 11) = Format(db(i, 11), "#,##0")
            End If
        Next
        
        Me.txtProductionTotalCost.Value = totalCost
        Me.txtProductionTotalCost.Text = Format(totalCost, "#,##0")
        
        Update_List Me.lstProductionList, db, "0pt;0pt;0pt,50pt,120pt;60pt;60pt;30pt;30pt;55pt;55pt;110pt;0pt"
        
    End If
    
    Me.txtProductionID.Value = ""
    
End Sub

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
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


Sub InsertProjection()
    Dim cost As Variant

    If Me.txtProductionItem.Value = "" Then MsgBox "품명을 입력하세요.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "금액을 입력하세요.": Exit Sub
    
    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    '예상실행항목에 저장
    Insert_Record shtProduction, CLng(Me.txtID.Value), Me.txtEstimateID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Date
    
    '예상실행가 계산
    Me.txtProductionTotalCost = GetProductionTotalCost
    
    '예상실행가 기준으로 비용 다시 계산
    CalculateEstimateUpdateCost
    
    '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "실행가", CLng(Me.txtProductionTotalCost.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "차액", CLng(Me.txtBidMargin.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "마진율", Me.txtBidMarginRate.Value
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "수주차액", CLng(Me.txtAcceptedMargin.Value)
    
    Me.txtProductionID.Value = ""
    
    '예상실행항목 리스트박스 새로고침
    InitializeLstProduction
    
End Sub


Sub UpdateProjection()
    Dim cost As Variant

    If Me.txtProductionID.Value = "" Then MsgBox "수정할 항목을 선택하세요.": Exit Sub
    
    If Me.txtProductionItem.Value = "" Then MsgBox "품명을 입력하세요.": Exit Sub
    If Me.txtProductionCost.Value = "" Then MsgBox "금액을 입력하세요.": Exit Sub

    If IsNumeric(Me.txtProductionCost.Value) Then
        cost = CLng(Me.txtProductionCost.Value)
    Else
        cost = Me.txtProductionCost.Value
    End If
    
    '기존 예상실행항목에 업데이트
    Update_Record shtProduction, Me.txtProductionID.Value, Me.txtID.Value, Me.txtEstimateID.Value, Me.txtProductionCustomer.Value, Me.txtProductionItem.Value, _
            Me.txtProductionMaterial.Value, Me.txtProductionSize.Value, _
            Me.txtProductionAmount.Value, Me.cboProductionUnit.Value, Me.txtProductionUnitPrice.Value, Me.txtProductionCost.Value, Me.txtProductionMemo.Value, Date
    
    '예상실행가 계산
    Me.txtProductionTotalCost = GetProductionTotalCost
    
    '예상실행가 기준으로 비용 다시 계산
    CalculateEstimateUpdateCost
    
    '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "실행가", CLng(Me.txtProductionTotalCost.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "차액", CLng(Me.txtBidMargin.Value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "마진율", Me.txtBidMarginRate.Value
    Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "수주차액", CLng(Me.txtAcceptedMargin.Value)
    
    InitializeLstProduction
    
    Select_ListItm Me.lstProductionList, Me.txtProductionID.Value
    
End Sub


Sub DeleteProjection()
    Dim db As Variant
    Dim YN As VbMsgBoxResult

    If Me.txtProductionID.Value = "" Then
        MsgBox "삭제할 항목을 선택하세요."
        Exit Sub
    Else
        '안내 문구 출력
        YN = MsgBox("선택한 항목을 삭제하시겠습니까? 삭제한 정보는 복구가 불가능합니다.", vbYesNo)
        If YN = vbNo Then Exit Sub
    
        '예상실행항목에서 삭제
        Delete_Record shtProduction, Me.txtProductionID.Value

        '예상실행가 계산
        Me.txtProductionTotalCost = GetProductionTotalCost
        
         '예상실행가 기준으로 비용 다시 계산
         CalculateEstimateUpdateCost
    
         '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
        Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "실행가", CLng(Me.txtProductionTotalCost.Value)
        Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "차액", CLng(Me.txtBidMargin.Value)
        Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "마진율", Me.txtBidMarginRate.Value
        Update_Record_Column shtEstimate, CLng(Me.txtID.Value), "수주차액", CLng(Me.txtAcceptedMargin.Value)
    
        Me.txtProductionID.Value = ""
    
        InitializeLstProduction
    
        ClearProductionInput
    End If
    
End Sub

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.Value, 2)
    
    'DB에 값이 있을 경우
    totalCost = 0
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 10)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(db(i, 10))
            End If
        Next
    End If
        
    GetProductionTotalCost = totalCost
End Function

'=============================================
'리스트박스 스크롤
'Private Sub lstProductionList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstProductionList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstProductionList
'End Sub






