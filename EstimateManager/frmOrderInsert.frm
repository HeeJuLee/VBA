VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderInsert 
   Caption         =   "발주 등록"
   ClientHeight    =   8775.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
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

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    'Label 위치 맞추기
    For Each contr In Me.Controls
    If contr.Name Like "Label*" Then
        contr.top = contr.top + 2
    End If
    Next
    
    '폼 위치 수정
    If orderInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = orderInsertFormX
        Me.top = orderInsertFormY
    End If
    
    InitializeOrderCategory
    InitializeCboUnit
    
    '발주명 입력창에 포커스
    Me.txtOrderName.SetFocus
    
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
            , , False
            
    Unload Me
    
    shtOrderAdmin.OrderSearch
    shtOrderAdmin.GoToEnd
    
End Sub


Function CheckOrderInsertValidation()
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
    
    CheckOrderInsertValidation = bCorrect
End Function

Sub CalculateOrderInsertCost()

    '수량값이 공백이면 발주금액은 단가
    If Me.txtAmount.value = "" Then
        Me.txtOrderPrice.value = Me.txtUnitPrice.value
        Exit Sub
    End If
    
    '단가와 수량을 곱한 값을 발주금액으로 세팅함
    If Me.txtUnitPrice.value <> "" And IsNumeric(Me.txtUnitPrice.value) Then
        Me.txtOrderPrice.value = CLng(Me.txtUnitPrice.value) * CLng(Me.txtAmount.value)
        Me.txtOrderPrice.Text = Format(Me.txtOrderPrice.value, "#,##0")
    End If

End Sub

Private Sub btnOrderClose_Click()
    Unload Me
    
    shtOrderAdmin.OrderSearch
End Sub

Private Sub btnOrderInsert_Click()
    InsertOrder
End Sub

Private Sub txtOrderName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub imgOrderDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    GetCalendarDate Me.txtOrderDate
End Sub

Private Sub txtOrderName_AfterUpdate()
    Me.lblOrderNameEmpty.Visible = False
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
    '오류메시지 숨김
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
    
    CalculateOrderInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblUnitPriceError.Visible = False
    
    If Me.txtUnitPrice.value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If IsNumeric(Me.txtUnitPrice.value) Then
            Me.txtUnitPrice.value = CLng(Me.txtUnitPrice.value)
        Else
            Me.txtUnitPrice.value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '금액 1,000자리 컴마 처리
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.value, "#,##0")
    End If
    
    CalculateOrderInsertCost
End Sub


Private Sub UserForm_Layout()
    orderInsertFormX = Me.Left
    orderInsertFormY = Me.top
End Sub


