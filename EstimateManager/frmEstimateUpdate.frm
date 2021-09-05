VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "견적 수정"
   ClientHeight    =   12240
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
    
    Me.txtProductionTotalCost.value = Format(estimate(17), "#,##0")   '실행가
    Me.txtBidPrice.value = Format(estimate(18), "#,##0")    '입찰가
    Me.txtBidMargin.value = Format(estimate(19), "#,##0")    '차액
    Me.txtBidMarginRate.value = Format(estimate(20), "0.0%")    '마진율
    Me.txtAcceptedPrice.value = Format(estimate(21), "#,##0")    '수주금액
    Me.txtAcceptedMargin.value = Format(estimate(22), "#,##0")   '수주차액
    
    Me.txtInsertDate.value = estimate(23)    '등록일자
    Me.txtUpdateDate.value = estimate(24)    '수정일자
    
    InitializeCboCategory
    Me.cboCategory.value = Trim(estimate(25))   '분류
    Me.txtDueDate.value = estimate(26)              '납기일
    Me.txtSpecificationDate.value = estimate(27)    '거래명세서
    Me.txtTaxInvoiceDate.value = estimate(28)    '세금계산서
    Me.txtPaymentDate.value = estimate(29)    '결제일자
    Me.txtExpectPaymentDate.value = estimate(30)  '예상결제일
    Me.txtExpectPaymentMonth.value = Format(estimate(30), "mm" & "월")  '예상결제월
    Me.txtVAT.value = Format(estimate(31), "#,##0")    '부가세
    Me.txtMemo.value = Trim(estimate(32))     '견적메모
    Me.chkVAT.value = estimate(33)      '부가세 제외 여부
    
    Me.txtPaid.value = Format(estimate(34), "#,##0")      '입금액
    Me.txtRemaining.value = Format(estimate(35), "#,##0")      '미입금액
    Me.chkDividePay.value = estimate(36)      '분할결제 여부
    If chkDividePay.value = True Then
        Me.btnPayment.Enabled = True
    Else
        Me.btnPayment.Enabled = False
    End If
    
    '예전에는 견적관리xls와 관리xls에서 메뉴가 다르게 남음
    '앞으로 견적메모와 수주메모는 같게 맞춰야 함.
    '견적메모와 수주메모가 다른 경우는 예전 경우임
    '(견적메모 = 견적메모 + 수주메모) 이렇게 맞추고 저장 시 수주쪽에도 동일하게 메모 넣을 예정
    acceptedMemo = Trim(estimate(37))
    If Me.txtMemo.value <> acceptedMemo Then
        If Me.txtMemo.value = "" Then
            Me.txtMemo.value = acceptedMemo
        Else
            Me.txtMemo.value = Me.txtMemo.value & vbCrLf & acceptedMemo
        End If
    End If
    
    '수주 ID (ID_관리)
    Me.txtAcceptedID.value = estimate(38)
    If Me.txtAcceptedID.value = "" Then
        '수주ID가 없으면 수주관련 컨트롤 unable 시킴
        frmOrder.Visible = False
        btnAcceptedInsert.Visible = True
        frmEstimateUpdate.Height = 280
    Else
        frmOrder.Visible = True
        btnAcceptedInsert.Visible = False
    End If
    
    '변경 전 관리번호
    orgManagementID = Me.txtManagementID
    
    InitializeLswOrderList      '발주 현황
    InitializeLswCustomerAutoComplete   '거래처 자동완성
    InitializeLswManagerAutoComplete    '담당자 자동완성
    
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

Sub UpdateEstimate()
    Dim db As Variant
    Dim blnUnique As Boolean
    Dim findRow As Long
    
    '입력 데이터 체크
    If CheckEstimateUpdateValidation = False Then
        Exit Sub
    End If

    '견적정보 DB 읽어오기
    db = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(db, Me.txtManagementID.value, 2, orgManagementID)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    '견적 테이블 업데이트
    Update_Record shtEstimate, Me.txtID.value, _
        Me.txtManagementID.value, Me.txtLinkedID.value, _
        Me.txtCustomer.value, Me.txtManager.value, _
        Me.txtEstimateName.value, Me.txtSize.value, _
        Me.txtAmount.value, Me.cboUnit.value, _
        Me.txtUnitPrice.value, Me.txtEstimatePrice.value, _
        Me.txtEstimateDate.value, Me.txtBidDate.value, _
        Me.txtAcceptedDate.value, Me.txtDeliveryDate.value, _
        Me.txtInsuranceDate.value, Me.txtProductionTotalCost.value, _
        Me.txtBidPrice.value, Me.txtBidMargin.value, _
        Me.txtBidMarginRate.value, Me.txtAcceptedPrice.value, _
        Me.txtAcceptedMargin.value, _
        Me.txtInsertDate.value, Date, _
        Me.cboCategory.value, Me.txtDueDate.value, _
        Me.txtSpecificationDate.value, Me.txtTaxInvoiceDate.value, Me.txtPaymentDate.value, Me.txtExpectPaymentDate.value, _
        Me.txtVAT.value, Me.txtMemo.value, Me.chkVAT.value, _
        Me.txtPaid.value, Me.txtRemaining.value, Me.chkDividePay, ""
    
    '수주 테이블 업데이트
    If Me.txtAcceptedID.value <> "" Then
        Update_Record shtOrder, Me.txtAcceptedID.value, _
        , Me.cboCategory.value, , _
        Me.txtManagementID.value, Me.txtCustomer.value, _
        Me.txtEstimateName.value, Me.txtManager.value, _
        Me.txtSize.value, Me.txtAmount.value, _
        Me.cboUnit.value, Me.txtUnitPrice, _
        Me.txtEstimatePrice.value, , _
        Me.txtAcceptedDate.value, , Me.txtDueDate.value, _
        , Me.txtDeliveryDate.value, _
        Me.txtSpecificationDate.value, Me.txtTaxInvoiceDate.value, Me.txtPaymentDate.value, Me.txtExpectPaymentDate.value, _
        , Me.txtVAT.value, _
        , Date, _
        , Me.txtMemo.value, Me.chkVAT.value
    End If
    
    '관리번호 변경이 되는 경우 대비하여 바꿔줌
    orgManagementID = Me.txtManagementID.value
    
    findRow = isExistInSheet(shtEstimateAdmin.Range("B6"), Me.txtID.value)
    If findRow <> 0 Then
        UpdateShtEstimate findRow
        
'        shtEstimateAdmin.Activate
'        shtEstimateAdmin.EstimateSearch
'        shtEstimateAdmin.Range("H" & findRow).Select
    End If
    
    findRow = isExistInSheet(shtOrderAdmin.Range("C6"), Me.txtID.value)
    If findRow <> 0 Then
        UpdateShtOrder findRow
'        shtOrderAdmin.Activate
'        shtOrderAdmin.OrderSearch
'        shtOrderAdmin.Range("I" & findRow).Select
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

    '예상차액과 예상마진율 계산
    If Me.txtBidPrice.value <> "" And Me.txtProductionTotalCost.value <> "" Then
        '예상차액 = 입찰가 - 예상실행가
        Me.txtBidMargin.value = Format(CLng(Me.txtBidPrice.value) - CLng(Me.txtProductionTotalCost.value), "#,##0")
        '예상마진율 = 예상차액 / 입찰가
        If Me.txtBidPrice.value <> "0" Then
            Me.txtBidMarginRate.value = Format(CLng(Me.txtBidMargin.value) / CLng(Me.txtBidPrice.value), "0.0%")
        End If
    Else
        Me.txtBidMargin.value = 0
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
    If Me.txtTaxInvoiceDate.value = "" Or chkVAT.value = True Then
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
            Me.txtUnitPrice.value, _
            Me.txtEstimatePrice.value, _
            , _
            , , , , , _
            , , , , _
            , , _
            Date, , _
            CLng(Me.txtID.value), , False

    '등록한 수주ID를 견적 테이블에 업데이트
    Update_Record_Column shtEstimate, Me.txtID, "ID_수주", Get_LastID(shtOrder)
    
    '폼을 새로 띄움
    Unload frmEstimateUpdate
    frmEstimateUpdate.Show (False)
    
End Sub

 Sub UpdateShtEstimate(findRow)
    shtEstimateAdmin.Cells(findRow, 4).value = Me.txtManagementID.value
    shtEstimateAdmin.Cells(findRow, 5).value = Me.txtCustomer.value
    shtEstimateAdmin.Cells(findRow, 6).value = Me.txtManager.value
    shtEstimateAdmin.Cells(findRow, 7).value = Me.cboCategory.value
    shtEstimateAdmin.Cells(findRow, 8).value = Me.txtEstimateName.value
    shtEstimateAdmin.Cells(findRow, 9).value = Me.txtSize.value
    shtEstimateAdmin.Cells(findRow, 10).value = Me.txtAmount.value
    shtEstimateAdmin.Cells(findRow, 11).value = Me.cboUnit.value
    shtEstimateAdmin.Cells(findRow, 12).value = Me.txtUnitPrice.value
    shtEstimateAdmin.Cells(findRow, 13).value = Me.txtEstimatePrice.value
    shtEstimateAdmin.Cells(findRow, 14).value = Me.txtEstimateDate.value
    shtEstimateAdmin.Cells(findRow, 15).value = Me.txtBidDate.value
    shtEstimateAdmin.Cells(findRow, 16).value = Me.txtAcceptedDate.value
    shtEstimateAdmin.Cells(findRow, 17).value = Me.txtDueDate.value
    shtEstimateAdmin.Cells(findRow, 18).value = Me.txtDeliveryDate.value
    shtEstimateAdmin.Cells(findRow, 19).value = Me.txtInsuranceDate.value
    shtEstimateAdmin.Cells(findRow, 20).value = Me.txtProductionTotalCost.value
    shtEstimateAdmin.Cells(findRow, 21).value = Me.txtBidPrice.value
    shtEstimateAdmin.Cells(findRow, 22).value = Me.txtBidMargin.value
    shtEstimateAdmin.Cells(findRow, 23).value = Me.txtBidMarginRate.value
    shtEstimateAdmin.Cells(findRow, 24).value = Me.txtAcceptedPrice.value
    shtEstimateAdmin.Cells(findRow, 25).value = Me.txtAcceptedMargin.value
    shtEstimateAdmin.Cells(findRow, 26).value = Me.txtSpecificationDate.value
    shtEstimateAdmin.Cells(findRow, 27).value = Me.txtTaxInvoiceDate.value
    shtEstimateAdmin.Cells(findRow, 28).value = Me.txtPaymentDate.value
    shtEstimateAdmin.Cells(findRow, 29).value = Me.txtExpectPaymentDate.value
    shtEstimateAdmin.Cells(findRow, 30).value = Me.txtVAT.value
    shtEstimateAdmin.Cells(findRow, 31).value = Me.txtInsertDate.value
    shtEstimateAdmin.Cells(findRow, 32).value = Date
End Sub

 Sub UpdateShtOrder(findRow)
    shtOrderAdmin.Cells(findRow, 5).value = Me.txtManagementID.value
    shtOrderAdmin.Cells(findRow, 6).value = Me.cboCategory.value
    shtOrderAdmin.Cells(findRow, 8).value = Me.txtCustomer.value
    shtOrderAdmin.Cells(findRow, 9).value = Me.txtEstimateName.value
    shtOrderAdmin.Cells(findRow, 10).value = Me.txtManager.value
    shtOrderAdmin.Cells(findRow, 11).value = Me.txtSize.value
    shtOrderAdmin.Cells(findRow, 12).value = Me.txtAmount.value
    shtOrderAdmin.Cells(findRow, 13).value = Me.cboUnit.value
    shtOrderAdmin.Cells(findRow, 14).value = Me.txtUnitPrice.value
    shtOrderAdmin.Cells(findRow, 15).value = Me.txtEstimatePrice.value
    shtOrderAdmin.Cells(findRow, 17).value = Me.txtAcceptedDate.value
    shtOrderAdmin.Cells(findRow, 19).value = Me.txtDueDate.value
    shtOrderAdmin.Cells(findRow, 21).value = Me.txtDeliveryDate.value
    shtOrderAdmin.Cells(findRow, 22).value = Me.txtSpecificationDate.value
    shtOrderAdmin.Cells(findRow, 23).value = Me.txtTaxInvoiceDate.value
    shtOrderAdmin.Cells(findRow, 24).value = Me.txtPaymentDate.value
    shtOrderAdmin.Cells(findRow, 25).value = Me.txtExpectPaymentDate.value
    shtOrderAdmin.Cells(findRow, 27).value = Me.txtVAT.value
    shtOrderAdmin.Cells(findRow, 28).value = Me.txtInsertDate.value
    shtOrderAdmin.Cells(findRow, 29).value = Date
End Sub

Function isExistInSheet(startRng As Range, value) As Long
    Dim WS As Worksheet
    Dim lastRow As Long
    Dim col As Long
    Dim i As Long
    Set WS = startRng.Parent
    
    lastRow = startRng.End(xlDown).row
    col = startRng.Column
    
    If IsNumeric(value) Then
        value = CLng(value)
    End If
    
    isExistInSheet = 0
    For i = startRng.row To lastRow
        If WS.Cells(i, col) = value Then
            isExistInSheet = i
            Exit Function
        End If
    Next
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

Private Sub btnProduction_Click()
    If isFormLoaded("frmProduction") Then
        Unload frmProduction
    End If
    frmProduction.Show (False)
End Sub


Private Sub btnAcceptedInsert_Click()
    InsertAccepted
End Sub

Private Sub btnPayment_Click()
    If isFormLoaded("frmPayment") Then
        Unload frmPayment
    End If
    frmPayment.Show (False)
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


Private Sub lswCustomerAutoComplete_DblClick()
    '거래처에 값을 넣어주고 포커스는 매니저로 이동
    With Me.lswCustomerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtCustomer.value = .SelectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
        End If
    End With
End Sub

Private Sub lswManagerAutoComplete_DblClick()
    '담당자명에 값을 넣어주고 포커스는 견적명으로 이동
    With Me.lswManagerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtManager.value = .SelectedItem.Text
            .Visible = False
            Me.txtEstimateName.SetFocus
        End If
    End With
End Sub

Private Sub txtManager_Enter()
    '자동완성 리스트에서 탭해서 넘어오는 경우
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .SelectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtEstimateName_Enter()
    '자동완성 리스트에서 탭해서 넘어오는 경우
    With Me.lswManagerAutoComplete
        If .Visible = True Then
            Me.txtManager.value = .SelectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = 13 Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtManager.SetFocus
        ElseIf KeyCode = 9 Then
            '탭키일 경우에 자동완성 결과가 하나이면 다음 입력칸으로 이동
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtManager.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = 40 Then
            '아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
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

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswManagerAutoComplete
        If KeyCode = 13 Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtEstimateName.SetFocus
        ElseIf KeyCode = 9 Then
            '탭키일 경우에 자동완성 결과가 하나이면 다음 입력칸으로 이동
            If .ListItems.count = 1 Then
                .Visible = False
                Me.txtEstimateName.SetFocus
                KeyCode = 0
            ElseIf .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        ElseIf KeyCode = 40 Then
            '아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End If
    End With
    
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

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처에 값을 넣어주고 포커스는 매니저로 이동
    If KeyCode = 13 Then
        With Me.lswCustomerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtCustomer.value = .SelectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
End Sub

Private Sub lswManagerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '담당자 선택 후 엔터키 들어오면 이 값을 담당자명에 넣어주고 포커스는 다음(사이즈)으로 이동
    If KeyCode = 13 Then
        With Me.lswManagerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtManager.value = .SelectedItem.Text
                .Visible = False
                Me.txtEstimateName.SetFocus
            End If
        End With
    End If
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
    
    Me.txtMemo.value = Trim(Me.txtMemo.value)
End Sub


Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateEstimateUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgExpectPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtExpectPaymentDate
    Me.txtExpectPaymentMonth = Format(Me.txtExpectPaymentDate, "mm" & "월")
End Sub

Private Sub txtManagementID_AfterUpdate()
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
End Sub

Private Sub txtEstimateName_AfterUpdate()
    Me.txtEstimateName.value = Trim(Me.txtEstimateName.value)
End Sub

Private Sub txtAmount_AfterUpdate()

    If Me.txtAmount.value <> "" Then
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '수량 1,000자리 컴마 처리
            Me.txtAmount.value = Format(Me.txtAmount.value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.value <> "" Then
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '견적단가 1,000자리 컴마 처리
            Me.txtUnitPrice.value = Format(Me.txtUnitPrice.value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtBidPrice_AfterUpdate()
    
    If Me.txtBidPrice.value <> "" Then
        If Not IsNumeric(Me.txtBidPrice.value) Then
            Me.txtBidPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '입찰금액 1,000자리 컴마 처리
            Me.txtBidPrice.value = Format(Me.txtBidPrice.value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtAcceptedPrice_AfterUpdate()
    If Me.txtAcceptedPrice.value <> "" Then
        If Not IsNumeric(Me.txtAcceptedPrice.value) Then
            Me.txtAcceptedPrice.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            Me.txtAcceptedPrice.value = Format(Me.txtAcceptedPrice.value, "#,##0")
            
            CalculateEstimateUpdateCost
        End If
    End If
End Sub

Private Sub txtProductionTotalCost_AfterUpdate()
    
    If Me.txtProductionTotalCost.value <> "" Then
        If Not IsNumeric(Me.txtProductionTotalCost.value) Then
            Me.txtProductionTotalCost.value = ""
            MsgBox "숫자를 입력하세요."
        Else
            '예상실행가 1,000자리 컴마 처리
            Me.txtProductionTotalCost.value = Format(Me.txtProductionTotalCost.value, "#,##0")
            
            '비용 필드 계산
            CalculateEstimateUpdateCost
        End If
    End If
    
End Sub

Private Sub txtExecutionCost_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub

Private Sub txtAcceptedDate_AfterUpdate()
    Me.txtAcceptedDate.value = Trim(Me.txtAcceptedDate.value)
    CalculateEstimateUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
    Me.txtTaxInvoiceDate.value = Trim(Me.txtTaxInvoiceDate.value)
   CalculateEstimateUpdateCost
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = Trim(Me.txtPaymentDate.value)
    CalculateEstimateUpdateCost
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


Private Sub cboUnit_AfterUpdate()
    Me.cboUnit.value = Trim(Me.cboUnit.value)
End Sub

Private Sub txtBidDate_AfterUpdate()
    Me.txtBidDate.value = Trim(Me.txtBidDate.value)
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
End Sub


Private Sub txtDeliveryDate_AfterUpdate()
    Me.txtDeliveryDate.value = Trim(Me.txtDeliveryDate.value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.value = Trim(Me.txtDueDate.value)
End Sub

Private Sub txtEstimateDate_AfterUpdate()
    Me.txtEstimateDate.value = Trim(Me.txtEstimateDate.value)
End Sub

Private Sub txtInsuranceDate_AfterUpdate()
    Me.txtInsuranceDate.value = Trim(Me.txtInsuranceDate.value)
End Sub

Private Sub txtManager_AfterUpdate()
    Me.txtManager.value = Trim(Me.txtManager.value)
End Sub


Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub


Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = Trim(Me.txtSpecificationDate.value)
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateEstimateUpdateCost
End Sub


Private Sub UserForm_Layout()
    estimateUpdateFormX = Me.Left
    estimateUpdateFormY = Me.top
End Sub



