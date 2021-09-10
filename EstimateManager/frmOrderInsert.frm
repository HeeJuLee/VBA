VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderInsert 
   Caption         =   "발주 등록"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   OleObjectBlob   =   "frmOrderInsert.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmOrderInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim bMatchedEstimateID As Boolean

Private Sub UserForm_Activate()
    '관리번호 입력창에 포커스
    Me.txtManagementID.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    '폼 위치 수정
    If orderInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderInsertFormX
        Me.top = orderInsertFormY
    End If
    
    '텍스트박스 라벨 위치 조정
    For Each contr In Me.Controls
        If contr.Name Like "txt*" Or contr.Name Like "cbo*" Or contr.Name Like "img*" Then
            contr.top = contr.top - 2
        End If
    Next
    
    InitializeCboUnit
    InitializeOrderCategory
    InitializeLswCustomerAutoComplete
    
    Me.txtOrderDate.value = Date
    bMatchedEstimateID = False
    
    
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

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 108
        .Visible = False
    End With
End Sub


Sub InsertOrder()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckOrderInsertValidation = False Then
        Exit Sub
    End If

    Insert_Record shtOrder, _
            , , Me.cboCategory.value, Me.txtManagementID.value, _
            Me.txtCustomer.value, _
            Me.txtOrderName.value, _
            Me.txtMaterial.value, _
            Me.txtSize.value, _
            Me.txtAmount.value, _
            Me.cboUnit.value, _
            Me.txtUnitPrice.value, _
            Me.txtOrderPrice.value, _
            Me.txtWeight.value, _
            , Me.txtOrderDate.value, , , , _
            , , , , _
            , , _
            Date, , _
            Me.txtEstimateID.value, , False
            
    Unload Me
    
    shtOrderAdmin.Activate
    shtOrderAdmin.OrderSearch
    shtOrderAdmin.GoToEnd
    
End Sub


Function CheckOrderInsertValidation()
    
    CheckOrderInsertValidation = False
    
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
    
    CheckOrderInsertValidation = True
End Function

Sub CalculateOrderInsertCost()

    '수량값이 공백이면 발주금액은 단가
    If Me.txtAmount.value = "" Then
        Me.txtOrderPrice.value = Me.txtUnitPrice.value
    Else
        If Me.txtUnitPrice.value = "" Then
            Me.txtOrderPrice.value = ""
        ElseIf IsNumeric(Me.txtUnitPrice.value) And IsNumeric(Me.txtAmount.value) Then
            Me.txtOrderPrice.value = Format(CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value), "#,##0")
        End If
    End If

End Sub

Private Sub btnOrderClose_Click()
    Unload Me
End Sub

Private Sub btnOrderInsert_Click()
    InsertOrder
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.lswCustomerAutoComplete
        If KeyCode = 13 Then
            '엔터키 - 다음 입력칸으로 이동
            .Visible = False
            Me.txtOrderName.SetFocus
        ElseIf KeyCode = 9 Then
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
        ElseIf KeyCode = 40 Then
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
    If KeyCode = 13 Then
        With Me.lswCustomerAutoComplete
            If Not .selectedItem Is Nothing Then
                Me.txtCustomer.value = .selectedItem.Text
                .Visible = False
                Me.txtOrderName.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub txtManagementID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
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
    
    CalculateOrderInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Me.txtUnitPrice.value = Trim(Me.txtUnitPrice.value)
    
    If Me.txtUnitPrice.value <> "" Then
        '단가값이 숫자가 아닐 경우 오류메시지 출력
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
    
    CalculateOrderInsertCost
End Sub

Private Sub txtOrderName_AfterUpdate()
    Me.txtOrderName.value = Trim(Me.txtOrderName.value)
End Sub

Private Sub txtCustomer_AfterUpdate()
    Me.txtCustomer.value = Trim(Me.txtCustomer.value)
End Sub

Private Sub txtMaterial_AfterUpdate()
    Me.txtMaterial.value = Trim(Me.txtMaterial.value)
End Sub

Private Sub txtOrderDate_AfterUpdate()
    Me.txtOrderDate.value = Trim(Me.txtOrderDate.value)
End Sub

Private Sub txtSize_AfterUpdate()
    Me.txtSize.value = Trim(Me.txtSize.value)
End Sub

Private Sub txtWeight_AfterUpdate()
    Me.txtWeight.value = Trim(Me.txtWeight.value)
End Sub


Private Sub UserForm_Layout()
    orderInsertFormX = Me.Left
    orderInsertFormY = Me.top
End Sub

Private Sub 발주정보_Click()

End Sub
