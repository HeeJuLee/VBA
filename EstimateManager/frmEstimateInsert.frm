VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "견적 등록"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   OleObjectBlob   =   "frmEstimateInsert.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEstimateInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Me.txtEstimateName.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim contr As Control
    
    'Label 위치 맞추기
'    For Each contr In Me.Controls
'    If contr.Name Like "Label*" Then
'        contr.top = contr.top + 2
'    End If
'    Next

    '컨트롤 초기화
    InitializeCboUnit
    InitializeLswCustomerAutoComplete
    InitializeLswManagerAutoComplete
    
    Me.txtEstimateDate.Value = Date
    
    '폼 위치값 조정
    If estimateInsertFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = estimateInsertFormX
        Me.top = estimateInsertFormY
    End If
    
    '텍스트박스 레이블 배경색 및 글자색 변경
    Me.lblEstimateName.BackColor = RGB(84, 130, 53)
    Me.lblManagementID.BackColor = RGB(84, 130, 53)
    Me.lblLinkedID.BackColor = RGB(48, 84, 150)
    Me.lblCustomer.BackColor = RGB(48, 84, 150)
    Me.lblManager.BackColor = RGB(48, 84, 150)
    Me.lblSize.BackColor = RGB(48, 84, 150)
    Me.lblAmount.BackColor = RGB(48, 84, 150)
    Me.lblUnit.BackColor = RGB(48, 84, 150)
    Me.lblUnitPrice.BackColor = RGB(48, 84, 150)
    Me.lblEstimatePrice.BackColor = RGB(48, 84, 150)
    Me.lblEstimateDate.BackColor = RGB(48, 84, 150)
    
    Me.lblEstimateName.ForeColor = RGB(255, 255, 255)
    Me.lblManagementID.ForeColor = RGB(255, 255, 255)
    Me.lblLinkedID.ForeColor = RGB(255, 255, 255)
    Me.lblCustomer.ForeColor = RGB(255, 255, 255)
    Me.lblManager.ForeColor = RGB(255, 255, 255)
    Me.lblSize.ForeColor = RGB(255, 255, 255)
    Me.lblAmount.ForeColor = RGB(255, 255, 255)
    Me.lblUnit.ForeColor = RGB(255, 255, 255)
    Me.lblUnitPrice.ForeColor = RGB(255, 255, 255)
    Me.lblEstimatePrice.ForeColor = RGB(255, 255, 255)
    Me.lblEstimateDate.ForeColor = RGB(255, 255, 255)
    
End Sub

Sub InitializeCboUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, db
End Sub

Sub InitializeLswCustomerAutoComplete()
    
    With Me.lswCustomerAutoComplete
        .View = lvwList
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

Sub InsertEstimate()
    Dim db As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckEstimateInsertValidation = False Then
        Exit Sub
    End If

    '견적정보 DB 읽어오기
    db = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(db, Me.txtManagementID.Value, 3)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    Insert_Record shtEstimate, _
            Trim(Me.txtManagementID.Value), _
            Trim(Me.txtLinkedID.Value), _
            Trim(Me.txtCustomer.Value), _
            Trim(Me.txtManager.Value), _
            Trim(Me.txtEstimateName.Value), _
            Trim(Me.txtSize.Value), _
            Trim(Me.txtAmount.Value), _
            Trim(Me.cboUnit.Value), _
            Trim(Me.txtUnitPrice.Value), _
            Trim(Me.txtEstimatePrice.Value), _
            Trim(Me.txtEstimateDate.Value), _
            , , , , _
            , , , , , , _
            Date, , _
            , , , , , , , , False
            
    Unload Me
    
    shtEstimateAdmin.Activate
    shtEstimateAdmin.EstimateSearch
    shtEstimateAdmin.GoToEnd
    
End Sub

Function CheckEstimateInsertValidation()

    '견적명이 입력되었는지 체크
    If Trim(Me.txtEstimateName.Value) = "" Then
        MsgBox "견적명을 입력하세요."
        CheckEstimateInsertValidation = False
        Me.txtEstimateName.SetFocus
        Exit Function
    End If
    
    '관리번호가 입력되었는지 체크
    If Trim(Me.txtManagementID.Value) = "" Then
        MsgBox "관리번호를 입력하세요."
        CheckEstimateInsertValidation = False
        Me.txtManagementID.SetFocus
        Exit Function
    End If
    
    CheckEstimateInsertValidation = True
End Function

Sub CalculateEstimateInsertCost()

    '수량값이 공백이면 견적금액은 견적단가
    If Me.txtAmount.Value = "" Then
        Me.txtEstimatePrice.Value = Me.txtUnitPrice.Value
        Exit Sub
    End If
    
    '견적단가와 수량을 곱한 값을 견적금액으로 세팅함
    If Me.txtUnitPrice.Value <> "" And IsNumeric(Me.txtUnitPrice.Value) Then
        Me.txtEstimatePrice.Value = CLng(Me.txtUnitPrice.Value) * CLng(Me.txtAmount.Value)
        Me.txtEstimatePrice.Text = Format(Me.txtEstimatePrice.Value, "#,##0")
    End If

End Sub

Private Sub btnEstimateClose_Click()
    Unload Me
End Sub

Private Sub btnEstimateInsert_Click()
    InsertEstimate
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub txtEstimateDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnEstimateInsert.SetFocus
    End If
End Sub

Private Sub txtCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '엔터키 - 다음 입력칸으로 이동
        Me.lswCustomerAutoComplete.Visible = False
        Me.txtManager.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '탭키, 아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
        With Me.lswCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtCustomer_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '거래처 자동완성 처리
    With Me.lswCustomerAutoComplete
        If Me.txtCustomer.Value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '견적거래처 DB를 읽어와서 리스트뷰에 출력
            .ListItems.Clear
            db = Get_DB(shtEstimateCustomer, True)
            db = Filtered_DB(db, Me.txtCustomer.Value, 1, False)
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

Private Sub lswCustomerAutoComplete_DblClick()
    '거래처에 값을 넣어주고 포커스는 품명으로 이동
    With Me.lswCustomerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtCustomer.Value = .SelectedItem.Text
            .Visible = False
            Me.txtManager.SetFocus
        End If
    End With
End Sub

Private Sub lswCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처 선택 후 엔터키 들어오면 이 값을 거래처명에 넣어주고 포커스는 다음(매니저명)으로 이동
    If KeyCode = 13 Then
        With Me.lswCustomerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtCustomer.Value = .SelectedItem.Text
                .Visible = False
                Me.txtManager.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManager_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '엔터키 - 다음 입력칸으로 이동
        Me.lswManagerAutoComplete.Visible = False
        Me.txtSize.SetFocus
    ElseIf KeyCode = 9 Or KeyCode = 40 Then
        '탭키, 아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
        With Me.lswManagerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .SelectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtManager_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim db As Variant
    Dim i As Long
    
    '담당자 자동완성 처리
    With Me.lswManagerAutoComplete
        If Me.txtManager.Value = "" Then
            .Visible = False
        Else
            .Visible = True
            
            '견적담당자 DB를 읽어와서 리스트뷰에 출력
            .ListItems.Clear
            db = Get_DB(shtEstimateManager, True)
            db = Filtered_DB(db, Me.txtManager.Value, 1, False)
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

Private Sub lswManagerAutoComplete_DblClick()
    '담당자명에 값을 넣어주고 포커스는 사이즈로 이동
    With Me.lswManagerAutoComplete
        If Not .SelectedItem Is Nothing Then
            Me.txtManager.Value = .SelectedItem.Text
            .Visible = False
            Me.txtSize.SetFocus
        End If
    End With
End Sub

Private Sub lswManagerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '담당자 선택 후 엔터키 들어오면 이 값을 담당자명에 넣어주고 포커스는 다음(사이즈)으로 이동
    If KeyCode = 13 Then
        With Me.lswManagerAutoComplete
            If Not .SelectedItem Is Nothing Then
                Me.txtManager.Value = .SelectedItem.Text
                .Visible = False
                Me.txtSize.SetFocus
            End If
        End With
    End If
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub txtAmount_AfterUpdate()
    
    If Me.txtAmount.Value <> "" Then
         '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.Value) Then
            MsgBox "숫자를 입력하세요."
            Me.txtAmount.Value = ""
            Exit Sub
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    CalculateEstimateInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    
    If Me.txtUnitPrice.Value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            MsgBox "숫자를 입력하세요."
            Me.txtUnitPrice.Value = ""
            Exit Sub
        End If
        
        '금액 1,000자리 컴마 처리
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    End If
    
    CalculateEstimateInsertCost
End Sub

Private Sub UserForm_Layout()
    estimateInsertFormX = Me.Left
    estimateInsertFormY = Me.top
End Sub
