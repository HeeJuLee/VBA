VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderUpdate 
   Caption         =   "발주 수정"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
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

Private Sub UserForm_Activate()
    Me.txtManagementID.SetFocus
End Sub

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
    If Not IsEmpty(db) Then
        '여러개 있을 경우에는 맨 마지막 견적정보 사용
        count = UBound(db, 1)
        Me.txtEstimateID.value = db(count, 1)
        Me.txtEstimateCustomer.value = db(count, 4)
        Me.txtEstimateManager.value = db(count, 5)
        Me.txtEstimateName.value = db(count, 6)
    
        bMatchedEstimateID = True
    End If
    
    InitializeCboUnit
    InitializeOrderPayMethod
    InitializeOrderCategory
    InitializeLswCustomerAutoComplete
    
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
    Me.txtReceivingDate.value = order(18)       '입고일자
    Me.txtSpecificationDate.value = order(20)   '명세서
    Me.txtTaxInvoiceDate.value = order(21)      '계산서
    Me.txtPaymentDate.value = order(22)     '결제일자
    Me.cboOrderPayMethod.value = Trim(order(24))       '결제수단
    Me.txtVAT.value = Format(order(25), "#,##0")             '부가세
    
    Me.txtInsertDate.value = order(26)    '등록일자
    Me.txtUpdateDate.value = order(27)    '수정일자
    
    Me.txtMemo.value = order(29)            '메모
    Me.chkVAT.value = order(30)             '부가세 제외 여부
    
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

Sub InitializeOrderPayMethod()
    Dim db As Variant
    db = Get_DB(shtOrderPayMethod, True)

    Update_Cbo Me.cboOrderPayMethod, db
End Sub

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 108
        .Visible = False
    End With
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
        Me.txtReceivingDate.value, , _
        Me.txtSpecificationDate.value, Me.txtTaxInvoiceDate.value, Me.txtPaymentDate.value, , _
        Me.cboOrderPayMethod.value, Me.txtVAT.value, _
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
    
    CheckOrderUpdateValidation = False
    
    '품목이 입력되었는지 체크
    If Trim(Me.txtOrderName.value) = "" Then
        MsgBox "품목을 입력하세요.", vbInformation, "작업 확인"
        Exit Function
    End If
    
    '관리번호가 입력되었고 유효한 관리번호인지 체크
    If Trim(Me.txtManagementID.value) = "" Then
        MsgBox "관리번호를 입력하세요.", vbInformation, "작업 확인"
        Exit Function
    End If
    
    If bMatchedEstimateID = False Then
        MsgBox "관리번호가 유효하지 않습니다.", vbInformation, "작업 확인"
        Exit Function
    End If
    
    CheckOrderUpdateValidation = True
    
End Function

Sub CalculateOrderUpdateCost()

    If Me.txtAmount.value = "" Then
        Me.txtOrderPrice.value = Me.txtUnitPrice.value
    Else
        If Me.txtUnitPrice.value = "" Then
            Me.txtOrderPrice.value = ""
        ElseIf IsNumeric(Me.txtUnitPrice.value) And IsNumeric(Me.txtAmount.value) Then
            Me.txtOrderPrice.value = Format(CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value), "#,##0")
        End If
    End If
    
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

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = vbKeyReturn Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtOrderName.SetFocus
        ElseIf KeyCode = vbKeyTab Then
            If .ListItems.count = 1 Then
                If Me.txtCustomer.value <> .ListItems(1).Text Then
                    '탭키일 경우 자동완성 결과와 입력값이 다르면 포커스를 자동완성 리스트로 이동
                    .selectedItem = .ListItems(1)
                    .SetFocus
                Else
                    '입력값과 자동완성 결과가 같으면 다음 입력칸으로 이동
                    .Visible = False
                    Me.txtOrderName.SetFocus
                End If
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
            
            '발주거래처 DB를 읽어와서 리스트뷰에 출력
            .ListItems.Clear
            db = Get_DB(shtOrderCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.value, 1, False)
            If IsEmpty(db) Then
                .Visible = False
            Else
                For i = 1 To UBound(db)
                    .ListItems.Add , , db(i, 1)
                    If i = 7 Then Exit For
                Next
            End If
            
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_DblClick()
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
    With Me.lswCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtOrderName.SetFocus
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처 선택 후 엔터키 들어오면 이 값을 거래처명에 넣어주고 포커스는 다음(품명)으로 이동
    If KeyCode = vbKeyReturn Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtOrderName.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtOrderName_Enter()
    '자동완성 리스트에서 탭해서 넘어오는 경우
    With Me.lswCustomerAutoComplete
        If .Visible = True Then
            Me.txtCustomer.value = .selectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub imgDueDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDueDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtReceivingDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxinvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
    CalculateOrderUpdateCost
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub txtManagementID_AfterUpdate()
    Dim db As Variant
    
    Me.txtManagementID.value = Trim(Me.txtManagementID.value)
    
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
            MsgBox "관리번호에 해당하는 견적(수주) 정보가 없습니다.", vbInformation, "작업 확인"
            Exit Sub
        Else
            If UBound(db, 1) = 1 Then
                Me.txtEstimateID.value = db(1, 1)
                Me.txtEstimateCustomer.value = db(1, 4)
                Me.txtEstimateManager.value = db(1, 5)
                Me.txtEstimateName.value = db(1, 6)
            
                bMatchedEstimateID = True
            Else
                MsgBox "관리번호가 여러개의 견적(수주) 정보에서 사용 중입니다.", vbInformation, "작업 확인"
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub txtAmount_AfterUpdate()
    Me.txtAmount.value = Trim(Me.txtAmount.value)
    
    If Me.txtAmount.value <> "" Then
         '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.value) Then
            Me.txtAmount.value = ""
            MsgBox "숫자를 입력하세요.", vbInformation, "작업 확인"
            Exit Sub
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Me.txtUnitPrice.value = Trim(Me.txtUnitPrice.value)
    
    If Me.txtUnitPrice.value <> "" Then
        '단가 값이 숫자가 아닐 경우 오류메시지 출력
        If IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = CLng(Me.txtUnitPrice.value)
        Else
            Me.txtUnitPrice.value = ""
            MsgBox "숫자를 입력하세요.", vbInformation, "작업 확인"
            Exit Sub
        End If
    End If
    
    '금액 1,000자리 컴마 처리
    Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    
    CalculateOrderUpdateCost
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
End Sub

Private Sub txtOrderName_AfterUpdate()
    Me.txtOrderName.value = Trim(Me.txtOrderName.value)
End Sub

Private Sub txtMaterial_AfterUpdate()
    Me.txtMaterial.value = Trim(Me.txtMaterial.value)
End Sub

Private Sub txtOrderDate_AfterUpdate()
    Me.txtOrderDate.value = ConvertDateFormat(Me.txtOrderDate.value)
End Sub

Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub

Private Sub txtWeight_AfterUpdate()
    Me.txtWeight.value = Trim(Me.txtWeight.value)
End Sub

Private Sub txtReceivingDate_AfterUpdate()
    Me.txtReceivingDate.value = ConvertDateFormat(Me.txtReceivingDate.value)
End Sub

Private Sub txtDueDate_AfterUpdate()
    Me.txtDueDate.value = ConvertDateFormat(Me.txtDueDate.value)
End Sub

Private Sub txtPaymentDate_AfterUpdate()
    Me.txtPaymentDate.value = ConvertDateFormat(Me.txtPaymentDate.value)
End Sub

Private Sub txtSpecificationDate_AfterUpdate()
    Me.txtSpecificationDate.value = ConvertDateFormat(Me.txtSpecificationDate.value)
End Sub

Private Sub txtTaxinvoiceDate_AfterUpdate()
    Me.txtTaxInvoiceDate.value = ConvertDateFormat(Me.txtTaxInvoiceDate.value)
   CalculateOrderUpdateCost
End Sub

Private Sub chkVAT_AfterUpdate()
    CalculateOrderUpdateCost
End Sub

Private Sub UserForm_Layout()
    orderUpdateFormX = Me.Left
    orderUpdateFormY = Me.top
End Sub

