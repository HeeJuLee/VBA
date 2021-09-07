VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate_1 
   Caption         =   "견적 수정"
   ClientHeight    =   13440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19125
   OleObjectBlob   =   "frmEstimateUpdate_1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEstimateUpdate_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim orgManagementID As Variant
Dim orgExecutionCost As String
Dim totlalCheckCount As Long




Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim estimateId As Variant
    Dim db As Variant
    Dim contr As Control
    
    If clickEstimateId <> "" Then              '발주관리에서 더블클릭한 경우
        If IsNumeric(clickEstimateId) Then
            estimateId = CLng(clickEstimateId)
        Else
            estimateId = clickEstimateId
        End If
        clickEstimateId = ""
    Else
        '선택한 행 번호
        cRow = Selection.row
    
        '데이터가 있는 행이 아닐 경우는 중지
        If cRow < 6 Or shtEstimateAdmin.Range("B" & cRow).value = "" Then
            MsgBox "수정할 견적 행을 먼저 선택한 후 견적수정 버튼을 클릭하세요."
            End
        End If
        
        estimateId = shtEstimateAdmin.Cells(cRow, 2)
    End If
    
    '텍스트박스 라벨 컨트롤 색상 조정
    For Each contr In Me.Controls
    If contr.Name Like "lbl*" Then
        'contr.top = contr.top + 2
        contr.BackColor = RGB(242, 242, 242)
    End If
    Next
    
    '폼 위치 수정
    If estimateUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateUpdateFormX
        Me.top = estimateUpdateFormY
    End If

    
    '견적 데이터 읽어오기
    estimate = Get_Record_Array(shtEstimate, estimateId)

    Me.txtID.value = estimate(1)    'ID
    Me.txtEstimateName.value = estimate(6)  '견적명
    Me.txtManagementID.value = estimate(2)    '관리번호
    Me.txtLinkedID.value = estimate(3)  '자재번호
    
    Me.txtCustomer = estimate(4)   '거래처
    Me.txtManager = estimate(5)   '담당자
    
    Me.txtSize.value = estimate(7)  '규격
    Me.txtAmount.value = Format(estimate(8), "#,##0")   '수량
    InitializeCboUnit
    Me.cboUnit.value = Trim(estimate(9))  '단위, ID가 없으므로 직접 value 넣으면 선택됨
    Me.txtUnitPrice.value = Format(estimate(10), "#,##0")     '견적단가
    Me.txtEstimatePrice.value = Format(estimate(11), "#,##0")     '견적금액
    
    Me.txtEstimateDate.value = estimate(12)    '견적일자
    Me.txtBidDate.value = estimate(13)    '입찰일자
    Me.txtAcceptedDate.value = estimate(14)    '수주일자
    Me.txtDeliveryDate.value = estimate(15)    '납품일자
    Me.txtInsuranceDate.value = estimate(16)    '증권일자
    
    Me.txtExecutionCost.value = Format(estimate(17), "#,##0")   '실행가
    orgExecutionCost = Me.txtExecutionCost.value
    Me.txtBidPrice.value = Format(estimate(18), "#,##0")    '입찰가
    Me.txtBidMargin.value = Format(estimate(19), "#,##0")    '차액
    Me.txtBidMarginRate.value = Format(estimate(20), "0.0%")    '마진율
    Me.txtAcceptedPrice.value = Format(estimate(21), "#,##0")    '수주금액
    Me.txtAcceptedMargin.value = Format(estimate(22), "#,##0")   '수주차액
    
    Me.txtInsertDate.value = estimate(23)    '등록일자
    Me.txtUpdateDate.value = estimate(24)    '수정일자
    
    InitializeCboCategory
    Me.cboCategory.value = Trim(estimate(25))   '분류
    '26은 납기일
    Me.txtSpecificationDate.value = estimate(27)    '거래명세서
    Me.txtTaxinvoiceDate.value = estimate(28)    '세금계산서
    Me.txtPaymentDate.value = estimate(29)    '결제일자
    Me.txtExpectPaymentDate.value = estimate(30)  '예상결제일
    Me.txtExpectPaymentMonth.value = Format(estimate(30), "mm" & "월")  '예상결제월
    Me.txtVAT.value = Format(estimate(31), "#,##0")    '부가세
    Me.txtMemo.value = estimate(32)
    Me.chkVAT.value = estimate(33)      '부가세 제외 여부
    
'    Me.txtExpectPay.Value = Format(estimate(27), "#,##0")    '입금예상액
'    Me.txtPaid.Value = Format(estimate(28), "#,##0")   '입금액
'    Me.txtUnpaid.Value = Format(estimate(29), "#,##0")   '미입금액
    
    '변경 전 관리번호
    orgManagementID = Me.txtManagementID
    
'    InitializeLswProductionList    '예상실행항목 목록
'    InitializeCboProductonUnit  '예상실행항목 단위
'    InitializeLswOrderCustomerAutoComplete   '발주거래처 자동완성
    
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
    db = Filtered_DB(db, Me.cboCustomer.value, 1, True)
    
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

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.value, 2, True)
    
     '리스트뷰 값 설정
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "품명", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "거래처", 50
        .ColumnHeaders.Add , , "재질", 60
        .ColumnHeaders.Add , , "규격", 60
        .ColumnHeaders.Add , , "수량", 30, lvwColumnRight
        .ColumnHeaders.Add , , "단위", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "단가", 60, lvwColumnRight
        .ColumnHeaders.Add , , "금액", 60, lvwColumnRight
        .ColumnHeaders.Add , , "메모", 110
        .ColumnHeaders.Add , , "등록일자", 0
        
        .CheckBoxes = True
        .ColumnHeaders(1).Position = 5
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 11)) Then
                    '비용 합계 구함
                    totalCost = totalCost + CLng(db(i, 11))
                End If
                
                Set li = .ListItems.Add(, , db(i, 5))
                li.ListSubItems.Add , , db(i, 1)
                li.ListSubItems.Add , , db(i, 2)
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , db(i, 8)
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                li.ListSubItems.Add , , db(i, 13)
            Next
            
            Me.txtProductionTotalCost.value = totalCost
            Me.txtProductionTotalCost.Text = Format(totalCost, "#,##0")
        End If
    End With
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

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
End Sub

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
    
    '견적ID에 해당하는 발주 정보를 읽어옴
    db = Get_DB(shtOrder)
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, Me.txtID.value, 28, True)
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
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "품명", 115
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "거래처", 50
        .ColumnHeaders.Add , , "재질", 60
        .ColumnHeaders.Add , , "규격", 60
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
        .ColumnHeaders.Add , , "결제수단", 59, lvwColumnCenter
        
        .ColumnHeaders(1).Position = 5
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 11)) Then
                    '비용 합계 구함
                    totalCost = totalCost + CLng(db(i, 11))
                End If
                
                Set li = .ListItems.Add(, , db(i, 7))   '품명
                li.ListSubItems.Add , , db(i, 1)        'ID
                li.ListSubItems.Add , , db(i, 28)       'ID_견적
                li.ListSubItems.Add , , db(i, 5)        '관리번호
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
                li.ListSubItems.Add , , db(i, 24)      '결제수단
            Next
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
    blnUnique = IsUnique(db, Me.txtManagementID.value, 3, orgManagementID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    '데이터 업데이트
    Update_Record shtEstimate, Me.txtID.value, _
        Me.txtManagementID.value, Me.txtLinkedID.value, _
        Me.txtCustomer.value, Me.txtManager.value, _
        Me.txtEstimateName.value, Me.txtSize.value, _
        Me.txtAmount.value, Me.cboUnit.value, _
        Me.txtUnitPrice.value, Me.txtEstimatePrice.value, _
        Me.txtEstimateDate.value, Me.txtBidDate.value, _
        Me.txtAcceptedDate.value, Me.txtDeliveryDate.value, _
        Me.txtInsuranceDate.value, Me.txtExecutionCost.value, _
        Me.txtBidPrice.value, Me.txtBidMargin.value, _
        Me.txtBidMarginRate.value, Me.txtAcceptedPrice.value, _
        Me.txtAcceptedMargin.value, _
        Me.txtInsertDate.value, Date, _
        Me.cboCategory.value, , _
        Me.txtSpecificationDate.value, Me.txtTaxinvoiceDate.value, Me.txtPaymentDate.value, Me.txtExpectPaymentDate.value, _
        Me.txtVAT.value, Me.txtMemo.value, Me.chkVAT.value
    
End Sub

Function CheckEstimateUpdateValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '견적명이 입력되었는지 체크
    If Me.txtEstimateName.value = "" Then
        bCorrect = False
        Me.lblErrorMessage.Caption = "견적명을 입력하세요."
    End If
    
    '관리번호가 입력되었는지 체크
    If Me.txtManagementID.value = "" Then
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
        If Me.txtAmount.value = "" Then
            Me.txtEstimatePrice.value = Me.txtUnitPrice.value
        Else
            Me.txtEstimatePrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.value, "#,##0")

    '차액과 마진율 계산
    If Me.txtBidPrice.value <> "" And Me.txtExecutionCost.value <> "" Then
        '차액 = 입찰가 - 실행가
        Me.txtBidMargin.value = CLng(Me.txtBidPrice.value) - CLng(Me.txtExecutionCost.value)
        Me.txtBidMargin.Text = Format(Me.txtBidMargin.value, "#,##0")
        '마진율 = 차액 / 입찰가
        Me.txtBidMarginRate.value = CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value)
        Me.txtBidMarginRate.Text = Format(Me.txtBidMarginRate.value, "0.0%")
    Else
        Me.txtBidMargin.value = 0
    End If

    '수주금액 계산
    If Me.txtAcceptedDate.value = "" Then
        '수주일자가 없는 경우
        Me.txtAcceptedPrice.value = 0
        Me.txtAcceptedMargin.value = 0
    Else
        '수주일자가 있는 경우
        '수주금액은 입찰금액으로 세팅
        If IsNumeric(Me.txtBidPrice.value) Then
            Me.txtAcceptedPrice.value = CLng(Me.txtBidPrice.value)
        Else
            Me.txtAcceptedPrice.value = 0
        End If
        Me.txtAcceptedPrice.Text = Format(Me.txtAcceptedPrice.value, "#,##0")
        
        '수주차액은 차액으로 세팅
        If IsNumeric(Me.txtBidMargin.value) Then
            Me.txtAcceptedMargin.value = CLng(Me.txtBidMargin.value)
        Else
            Me.txtAcceptedMargin.value = 0
        End If
        Me.txtAcceptedMargin.Text = Format(Me.txtAcceptedMargin.value, "#,##0")
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


Sub InsertProduction()
    Dim cost As Variant

    If Me.txtProductionItem.value = "" Then MsgBox "품명을 입력하세요.": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "금액을 입력하세요.": Exit Sub
    
    If IsNumeric(Me.txtProductionCost.value) Then
        cost = CLng(Me.txtProductionCost.value)
    Else
        cost = Me.txtProductionCost.value
    End If
    
    '예상실행항목에 저장
    Insert_Record shtProduction, CLng(Me.txtID.value), Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Date
    
    '예상실행항목 합계 계산
    Me.txtProductionTotalCost.value = GetProductionTotalCost
    Me.txtExecutionCost.value = Me.txtProductionTotalCost.value

    '실행가 기준으로 비용 다시 계산
    CalculateEstimateUpdateCost
    
    '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "실행가", CLng(Me.txtProductionTotalCost.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "차액", CLng(Me.txtBidMargin.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "마진율", Me.txtBidMarginRate.value
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "수주차액", CLng(Me.txtAcceptedMargin.value)
    
    '예상실행항목 리스트박스 새로고침
    InitializeLswProductionList
    
    '등록한 아이템 선택
    Me.txtProductionID.value = Get_LastID(shtProduction)
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub


Sub UpdateProduction()
    Dim cost As Variant

    If Me.txtProductionID.value = "" Then MsgBox "수정할 항목을 선택하세요.": Exit Sub
    
    If Me.txtProductionItem.value = "" Then MsgBox "품명을 입력하세요.": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "금액을 입력하세요.": Exit Sub

    If IsNumeric(Me.txtProductionCost.value) Then
        cost = CLng(Me.txtProductionCost.value)
    Else
        cost = Me.txtProductionCost.value
    End If
    
    '기존 예상실행항목에 업데이트
    Update_Record shtProduction, Me.txtProductionID.value, Me.txtID.value, Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Date
    
    '예상실행가 계산
    Me.txtProductionTotalCost.value = GetProductionTotalCost
    Me.txtExecutionCost.value = Me.txtProductionTotalCost.value
    
    '예상실행가 기준으로 비용 다시 계산
    CalculateEstimateUpdateCost
    
    '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "실행가", CLng(Me.txtProductionTotalCost.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "차액", CLng(Me.txtBidMargin.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "마진율", Me.txtBidMarginRate.value
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "수주차액", CLng(Me.txtAcceptedMargin.value)
    
    InitializeLswProductionList
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub


Sub DeleteProduction()
    Dim db As Variant
    Dim YN As VbMsgBoxResult

    If Me.txtProductionID.value = "" Then MsgBox "삭제할 항목을 선택하세요.": Exit Sub
        
    '안내 문구 출력
    YN = MsgBox("선택한 항목을 삭제하시겠습니까? 삭제한 정보는 복구가 불가능합니다.", vbYesNo)
    If YN = vbNo Then Exit Sub

    '예상실행항목에서 삭제
    Delete_Record shtProduction, Me.txtProductionID.value

    '예상실행가 계산
    Me.txtProductionTotalCost.value = GetProductionTotalCost
    Me.txtExecutionCost.value = Me.txtProductionTotalCost.value
    
     '예상실행가 기준으로 비용 다시 계산
     CalculateEstimateUpdateCost

     '예상실행가, 입찰차액, 마진율, 수주차액 금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "실행가", CLng(Me.txtProductionTotalCost.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "차액", CLng(Me.txtBidMargin.value)
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "마진율", Me.txtBidMarginRate.value
    Update_Record_Column shtEstimate, CLng(Me.txtID.value), "수주차액", CLng(Me.txtAcceptedMargin.value)

    Me.txtProductionID.value = ""

    InitializeLswProductionList

    ClearProductionInput
    
End Sub

Sub ProductionToOrder()
    Dim li As ListItem
    Dim count As Long
    Dim managementId, customer, Item, material, size, amount, unit, unitPrice, cost, memo As Variant
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Checked = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "발주할 항목을 체크박스에 체크하세요.": Exit Sub
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Checked = True Then
            Item = li.Text
            managementId = li.SubItems(3)
            customer = li.SubItems(4)
            material = li.SubItems(5)
            size = li.SubItems(6)
            amount = li.SubItems(7)
            unit = li.SubItems(8)
            unitPrice = li.SubItems(9)
            cost = li.SubItems(10)
            memo = li.SubItems(11)
            
            '선택한 예상실행항목을 발주 테이블에 등록
            Insert_Record shtOrder, _
                , , managementId, customer, Item, material, size, amount, unit, unitPrice, cost, _
                , , , , , _
                , , , , , _
                Date, , Me.txtID, False, memo
                
            count = count + 1
        End If
    Next
    
    InitializeLswOrderList
    
    MsgBox "총 " & count & "개 항목을 발주하였습니다.", vbInformation
    
    shtOrderAdmin.Activate
    shtOrderAdmin.OrderSearch
    shtOrderAdmin.GoToEnd

End Sub

Function GetProductionTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, Me.txtID.value, 2, True)
    
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

Sub SelectItemLswProduction(selectedID As Variant)
    Dim i As Long
    
    With Me.lswProductionList
        If Not IsMissing(selectedID) Then
            For i = 1 To .ListItems.count
                If selectedID = .ListItems(i).SubItems(1) Then
                    .selectedItem = .ListItems(i)
                    .SetFocus
                End If
            Next
        End If
    End With
End Sub

Private Sub lswProductionList_Click()
    With Me.lswProductionList
        If Not .selectedItem Is Nothing Then
            Me.txtProductionID.value = .selectedItem.ListSubItems(1)
            Me.txtProductionItem.value = .selectedItem.Text
            Me.txtProductionCustomer.value = .selectedItem.ListSubItems(4)
            Me.txtProductionMaterial.value = .selectedItem.ListSubItems(5)
            Me.txtProductionSize.value = .selectedItem.ListSubItems(6)
            Me.txtProductionAmount.value = .selectedItem.ListSubItems(7)
            Me.cboProductionUnit.value = .selectedItem.ListSubItems(8)
            Me.txtProductionUnitPrice.value = .selectedItem.ListSubItems(9)
            Me.txtProductionCost.value = .selectedItem.ListSubItems(10)
            Me.txtProductionMemo.value = .selectedItem.ListSubItems(11)
        End If
    End With
End Sub

Private Sub lswOrderList_DblClick()
    With Me.lswOrderList
        If Not .selectedItem Is Nothing Then
            clickOrderId = .selectedItem.ListSubItems(1)
            
            If frmOrderUpdate.Visible = True Then
                Unload frmOrderUpdate
            End If
        
            frmOrderUpdate.Show (False)
        End If
    End With
End Sub

Private Sub btnEstimateUpdate_Click()
`    UpdateEstimate
    
    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.Range("H" & selectionRow).Select
End Sub

Private Sub btnEstimateClose_Click()
    If orgExecutionCost <> Me.txtExecutionCost.value Then
        Unload Me
        
        '실행가가 변경된 경우에만 견적관리 화면 새로고침
        shtEstimateAdmin.Activate
        shtEstimateAdmin.EstimateSearch
    Else
        Unload Me
    End If
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

Private Sub lswProductionList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    With Me.lswProductionList
        .selectedItem.Selected = False
        
        If Item.Checked = True Then
            Item.Bold = True
            Item.ForeColor = vbBlue
            totlalCheckCount = totlalCheckCount + 1
        Else
            Item.Bold = False
            Item.ForeColor = vbBlack
            totlalCheckCount = totlalCheckCount - 1
        End If
    End With
    
    If totlalCheckCount = 0 Then
        Me.btnProductionToOrder.Caption = "체크 항목 발주"
    Else
        Me.btnProductionToOrder.Caption = totlalCheckCount & "개 항목 발주"
    End If
End Sub


Private Sub txtProductionCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '엔터키 - 다음 입력칸으로 이동
        Me.lswOrderCustomerAutoComplete.Visible = False
        Me.txtProductionItem.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '탭키, 아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtProductionCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '거래처 자동완성 처리
    With Me.lswOrderCustomerAutoComplete
        If Me.txtProductionCustomer.value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '발주거래처 DB를 읽어와서 리스트뷰에 출력
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
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
    With Me.lswOrderCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtProductionCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtProductionItem.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
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

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxinvoiceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "월")
End Sub

Private Sub cboCustomer_Change()
    InitializeCboManager
End Sub

Private Sub txtEstimateDate_Change()
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.lblErrorMessage.Visible = False
End Sub

Private Sub txtAmount_AfterUpdate()
    '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    '수량값이 공백이면 중지
    If Me.txtAmount.value <> "" Then
        '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtUnitPrice.value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If
    
    '견적단가 1,000자리 컴마 처리
    Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
End Sub

Private Sub txtBidPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblErrorMessage.Visible = False
    
    If Me.txtBidPrice.value <> "" Then
        '입찰금액이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtBidPrice.value) Then
            Me.txtBidPrice.value = ""
            Me.lblErrorMessage.Caption = "숫자를 입력하세요."
            Me.lblErrorMessage.Visible = True
        End If
    End If

    '입찰금액 1,000자리 컴마 처리
    Me.txtBidPrice.Text = Format(Me.txtBidPrice.value, "#,##0")
    
    '비용 필드 계산
    CalculateEstimateUpdateCost
    
End Sub

Private Sub txtProductionAmount_AfterUpdate()
    
    If Me.txtProductionAmount.value = "" Then
        Exit Sub
    End If
    
    If IsNumeric(Me.txtProductionAmount.value) Then
        Me.txtProductionAmount.Text = Format(Me.txtProductionAmount.value, "#,##0")
        
        '금액 = 수량 * 단가
        If IsNumeric(Me.txtProductionUnitPrice.value) Then
            Me.txtProductionCost.value = CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value)
            Me.txtProductionCost.Text = Format(Me.txtProductionCost.value, "#,##0")
        End If
    End If
End Sub

Private Sub txtProductionUnitPrice_AfterUpdate()
    
    If Me.txtProductionUnitPrice.value = "" Then
        Exit Sub
    End If
    
    If IsNumeric(Me.txtProductionUnitPrice.value) Then
        Me.txtProductionUnitPrice.Text = Format(Me.txtProductionUnitPrice.value, "#,##0")
        
        If Me.txtProductionAmount.value = "" Then
            Me.txtProductionCost.value = Me.txtProductionUnitPrice.value
            Me.txtProductionCost.Text = Format(Me.txtProductionCost.value, "#,##0")
        Else
            If IsNumeric(Me.txtProductionAmount.value) Then
                '금액 = 수량 * 단가
                Me.txtProductionCost.value = CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value)
                Me.txtProductionCost.Text = Format(Me.txtProductionCost.value, "#,##0")
            End If
        End If
    End If
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub

