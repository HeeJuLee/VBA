VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderUpdate 
   Caption         =   "발주 수정"
   ClientHeight    =   9435.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   OleObjectBlob   =   "frmOrderUpdate.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmOrderUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bMatchedEstimateID As Boolean

Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim order As Variant
    Dim db As Variant
    Dim contr As Control
    Dim orderId As Long
    Dim pos As Long
    Dim count As Long
    
    If clickOrderId <> "" Then              '견적수정 폼의 발주현황에서 더블클릭한 경우
        If IsNumeric(clickOrderId) Then
            orderId = CLng(clickOrderId)
        Else
            orderId = clickOrderId
        End If
        clickOrderId = ""
    Else
        cRow = Selection.row                '발주관리화면에서 더블클릭으로 선택한 행 번호

        If cRow < 6 Or shtOrderAdmin.Range("B" & cRow).value = "" Then End         '데이터가 있는 행이 아닐 경우는 중지
        
        orderId = shtOrderAdmin.Cells(cRow, 2)
    End If
    
    'Label 위치 맞추기
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '폼 위치 수정
    If orderUpdateFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderUpdateFormX
        Me.top = orderUpdateFormY
    End If
    
    '발주 데이터 읽어오기
    order = Get_Record_Array(shtOrder, orderId)
    
    Me.txtID.value = order(1)   'ID
    Me.txtManagementID.value = order(5) '관리번호
    
    '관리번호로 견적정보 가져오기
    bMatchedEstimateID = False
    db = Get_DB(shtEstimate)
    db = Filtered_DB(db, Me.txtManagementID.value, 2, True)
    If IsEmpty(db) Then
        Me.lblManagementIDError.Caption = "관리번호 오류"
        Me.lblManagementIDError.Visible = True
    Else
        '여러개 있을 경우에는 맨 마지막 견적정보 사용
        count = UBound(db, 1)
        Me.txtEstimateID.value = db(count, 1)
        Me.txtEstimateCustomer.value = db(count, 4)
        Me.txtEstimateManager.value = db(count, 5)
        Me.txtEstimateName.value = db(count, 6)
    
        bMatchedEstimateID = True
    End If
    
    InitializeOrderCategory
    InitializeCboUnit
    InitializePayMethod
    
    Me.cboCategory.value = Trim(order(4))     '분류
    Me.txtCustomer.value = order(6)     '거래처
    Me.txtOrderName.value = order(7)    '발주 품명
    Me.txtMaterial.value = order(8)     '재질
    Me.txtSize.value = order(9)             '규격
    Me.txtAmount.value = Format(order(10), "#,##0")   '수량
    Me.cboUnit.value = Trim(order(11))            '단위
    Me.txtUnitPrice.value = Format(order(12), "#,##0")     '단가
    Me.txtOrderPrice.value = Format(order(13), "#,##0")      '발주금액
    Me.txtWeight.value = order(14)          '중량
    Me.txtOrderDate.value = order(16)       '발주일자
    Me.txtDueDate.value = order(17)         '납기일자
    Me.txtDeliveryDate.value = order(18)       '입고일자
    Me.txtSpecificationDate.value = order(20)   '명세서
    Me.txtTaxInvoiceDate.value = order(21)      '계산서
    Me.txtPaymentDate.value = order(22)     '결제일자
    Me.cboPayMethod.value = Trim(order(24))       '결제수단
    Me.txtVAT.value = Format(order(25), "#,##0")             '부가세
    
    Me.txtInsertDate.value = order(26)    '등록일자
    Me.txtUpdateDate.value = order(27)    '수정일자
    
    Me.txtMemo.value = order(29)            '메모
    Me.chkVAT.value = order(30)             '부가세 제외 여부
    
    '발주명 입력창에 포커스
    Me.txtOrderName.SetFocus
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeOrderCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboCategory, db
End Sub

Sub InitializePayMethod()
    Dim db As Variant
    db = Get_DB(shtPayMethod, True)

    Update_Cbo Me.cboPayMethod, db
End Sub


Sub UpdateOrder()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckOrderUpdateValidation = False Then
        Exit Sub
    End If

    '데이터 업데이트
    Update_Record shtOrder, Me.txtID.value, _
        , , Me.cboCategory.value, _
        Me.txtManagementID.value, Me.txtCustomer.value, _
        Me.txtOrderName.value, Me.txtMaterial.value, _
        Me.txtSize.value, Me.txtAmount.value, _
        Me.cboUnit.value, Me.txtUnitPrice, _
        Me.txtOrderPrice.value, Me.txtWeight.value, _
        , Me.txtOrderDate.value, Me.txtDueDate.value, _
        Me.txtDeliveryDate.value, , _
        Me.txtSpecificationDate.value, Me.txtTaxInvoiceDate.value, Me.txtPaymentDate.value, , _
        Me.cboPayMethod.value, Me.txtVAT.value, _
        Me.txtInsertDate, Date, _
        Me.txtEstimateID.value, Me.txtMemo.value, Me.chkVAT.value

    Unload Me
    
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.InitializeLswOrderList
    Else
        shtOrderAdmin.Activate
        shtOrderAdmin.OrderSearch
        shtOrderAdmin.Range("K" & selectionRow).Select
    End If
    
End Sub


Function CheckOrderUpdateValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '발주명이 입력되었는지 체크
    If Trim(Me.txtOrderName.value) = "" Then
        bCorrect = False
        Me.lblOrderNameEmpty.Visible = True
    Else
        Me.lblOrderNameEmpty.Visible = False
    End If
    
    '관리번호가 입력되었고 유효한 관리번호인지 체크
    If Trim(Me.txtManagementID.value) = "" Or bMatchedEstimateID = False Then
        bCorrect = False
        Me.lblManagementIDEmpty.Visible = True
    Else
        Me.lblManagementIDEmpty.Visible = False
    End If
    
    CheckOrderUpdateValidation = bCorrect
End Function

Sub CalculateOrderUpdateCost()

    '발주금액 계산
    If Me.txtUnitPrice.value <> "" Then
        If Me.txtAmount.value = "" Then
            Me.txtOrderPrice.value = Me.txtUnitPrice.value
        Else
            Me.txtOrderPrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        End If
    End If
    Me.txtOrderPrice.Text = Format(Me.txtOrderPrice.value, "#,##0")
    
    '부가세 계산
    '세금계산서 일자가 없는 경우, 부가세 제외인 경우 부가세는 0
    If Me.txtTaxInvoiceDate.value = "" Or chkVAT.value = True Then
        Me.txtVAT.value = 0
    Else
        '부가세는 금액의 10%
        If Me.txtOrderPrice.value <> "" And Me.txtOrderPrice.value <> 0 Then
            Me.txtVAT.value = CLng(Me.txtOrderPrice.value) * 0.1
            Me.txtVAT.Text = Format(Me.txtVAT.value, "#,##0")
        End If
    End If

End Sub

Private Sub btnOrderUpdate_Click()
    UpdateOrder
End Sub

Private Sub btnOrderClose_Click()
    Unload Me
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateOrderUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub txtManagementID_AfterUpdate()
    Dim db As Variant
    
    Me.lblManagementIDEmpty.Visible = False
    Me.lblManagementIDError.Visible = False
    
    Me.txtEstimateID.value = ""
    Me.txtEstimateCustomer.value = ""
    Me.txtEstimateManager.value = ""
    Me.txtEstimateName.value = ""
    
    '입력한 관리번호로 견적테이블을 검색해서 견적ID를 가져옴
    bMatchedEstimateID = False
    If Me.txtManagementID.value <> "" Then
        db = Get_DB(shtEstimate)
        db = Filtered_DB(db, Me.txtManagementID.value, 2, True)
        If IsEmpty(db) Then
            Me.lblManagementIDError.Caption = "관리번호 오류"
            Me.lblManagementIDError.Visible = True
        Else
            If UBound(db, 1) = 1 Then
                Me.txtEstimateID.value = db(1, 1)
                Me.txtEstimateCustomer.value = db(1, 4)
                Me.txtEstimateManager.value = db(1, 5)
                Me.txtEstimateName.value = db(1, 6)
            
                bMatchedEstimateID = True
            Else
                Me.lblManagementIDError.Caption = "관리번호 중복"
                Me.lblManagementIDError.Visible = True
            End If
        End If
    End If
    
End Sub

Private Sub txtAmount_AfterUpdate()
    Me.lblAmountError.Visible = False
    
    If Me.txtAmount.value <> "" Then
         '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            Me.lblAmountError.Visible = True
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Me.lblUnitPriceError.Visible = False
    
    If Me.txtUnitPrice.value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '단가 1,000자리 컴마 처리
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    End If
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtTaxInvoiceDate_AfterUpdate()
   CalculateOrderUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateOrderUpdateCost
End Sub

Private Sub UserForm_Layout()
    orderUpdateFormX = Me.Left
    orderUpdateFormY = Me.top
End Sub

