VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateInsert 
   Caption         =   "견적 등록"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "frmEstimateInsert.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEstimateInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub txtEstimateDate_Change()
    '오류 메시지 숨김
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateID_AfterUpdate()
    '오류 메시지 숨김
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateName_AfterUpdate()
    '오류 메시지 숨김
    Me.lblInputFieldError.Visible = False
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub UserForm_Initialize()

    '거래처, 담당자 콤보박스 세팅
    InitializeCboCustomer
    InitializeCboUnit
    
    '견적명 입력창에 포커스
    Me.txtEstimateName.SetFocus
    
End Sub

Private Sub cboCustomer_Change()
    '콤보박스에서 거래처를 변경하면 해당 거래처의 담당자로 담당자 콤보박스를 세팅
    InitializeCboManager
End Sub

Private Sub txtAmount_AfterUpdate()
    '오류메시지 숨김
    Me.lblAmountError.Visible = False
    Me.lblInputFieldError.Visible = False
    
    If Me.txtAmount.Value <> "" Then
         '수량값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtAmount.Value) Then
            Me.txtAmount.Value = ""
            Me.lblAmountError.Visible = True
        End If
    End If
    
    '수량 1,000자리 컴마 처리
    Me.txtAmount.Text = Format(Me.txtAmount.Value, "#,##0")
    
    CalculateEstimateInsertCost
End Sub

Private Sub txtUnitPrice_AfterUpdate()
     '오류메시지 숨김
    Me.lblUnitPriceError.Visible = False
    Me.lblInputFieldError.Visible = False
    
    If Me.txtUnitPrice.Value <> "" Then
        '견적단가값이 숫자가 아닐 경우 오류메시지 출력
        If Not IsNumeric(Me.txtUnitPrice.Value) Then
            Me.txtUnitPrice.Value = ""
            Me.lblUnitPriceError.Visible = True
            Exit Sub
        End If
        
        '금액 1,000자리 컴마 처리
        Me.txtUnitPrice.Text = Format(Me.txtUnitPrice.Value, "#,##0")
    End If
    
    CalculateEstimateInsertCost
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub


Private Sub btnEstimateInsert_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    InsertEstimate
End Sub

Private Sub btnEstimateClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub

'버튼 마우스오버 처리
'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnEstimateInsert_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEstimateInsert
End Sub

Private Sub btnEstimateInsert_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEstimateInsert
End Sub

Private Sub btnEstimateInsert_Enter()
OnHover_Css Me.btnEstimateInsert
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnEstimateClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEstimateClose
End Sub

Private Sub btnEstimateClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEstimateClose
End Sub

Private Sub btnEstimateClose_Enter()
OnHover_Css Me.btnEstimateClose
End Sub

'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnEstimateInsert, btnEstimateClose" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
 For Each vList In vLists
 If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
 Next
Next
End Sub
'커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
Private Sub OnHover_Css(lbl As Control): With lbl: .BackColor = RGB(211, 240, 224): .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl: .BackColor = &H8000000E: .BorderColor = -2147483638: End With: End Sub


Sub InitializeCboUnit()
    Dim DB As Variant
    DB = Get_DB(shtUnit, True)

    Update_Cbo Me.cboUnit, DB
End Sub

Sub InitializeCboCustomer()
    Dim DB As Variant
    DB = Get_DB(shtCustomer)

    Update_Cbo Me.cboCustomer, DB, 2
End Sub

Sub InitializeCboManager()
    Dim DB As Variant
    
    '담당자 DB를 읽어와서
    DB = Get_DB(shtManager)
    '거래처ID로 필터링
    DB = Filtered_DB(DB, Me.cboCustomer.Value, 2)
    
    '기존 콤보박스 내용지우기
    Me.cboManager.Clear
    
    '담당자가 있으면 콤보박스에 추가함
    If Not IsEmpty(DB) Then
        Update_Cbo Me.cboManager, DB, 3
    End If
End Sub



Sub InsertEstimate()
    Dim DB As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
    If CheckEstimateInsertValidation = False Then
        Exit Sub
    End If

    '견적정보 DB 읽어오기
    DB = Get_DB(shtEstimate)
    
    '동일한 관리번호가 있는지 체크
    blnUnique = IsUnique(DB, Me.txtEstimateID.Value, 3)
    If blnUnique = False Then MsgBox "동일한 관리번호가 존재합니다. 다시 확인해주세요.", vbExclamation: Exit Sub
    
    Insert_Record shtEstimate, _
            Me.cboManager.Value, _
            Me.txtEstimateID.Value, _
            Me.txtLinkedID.Value, _
            Me.txtEstimateName.Value, _
            Me.txtSize.Value, _
            Me.txtAmount.Value, _
            Me.cboUnit.Value, _
            Me.txtUnitPrice.Value, _
            Me.txtEstimatePrice.Value, _
            Me.txtEstimateDate.Value, _
            , , , , _
            , , , , , , _
            , , , , _
            , , , , _
            Date, Date

    Unload Me
    
    shtEstimateAdmin.EstimateSearch
    
End Sub


Function CheckEstimateInsertValidation()
    Dim bCorrect As Boolean
    
    bCorrect = True
    
    '견적명이 입력되었는지 체크
    If Trim(Me.txtEstimateName.Value) = "" Then
        bCorrect = False
        Me.lblEstimateNameEmpty.Visible = True
    Else
        Me.lblEstimateNameEmpty.Visible = False
    End If
    
    '관리번호가 입력되었는지 체크
    If Trim(Me.txtEstimateID.Value) = "" Then
        bCorrect = False
        Me.lblEstimateIDEmpty.Visible = True
    Else
        Me.lblEstimateIDEmpty.Visible = False
    End If
    
    '수량이 입력되었는지 체크
    If Trim(Me.txtAmount.Value) = "" Then
        bCorrect = False
        Me.lblAmountEmpty.Visible = True
    Else
        Me.lblAmountEmpty.Visible = False
    End If
    
    '견적단가가 입력되었는지 체크
    If Trim(Me.txtUnitPrice.Value) = "" Then
        bCorrect = False
        Me.lblUnitPriceEmpty.Visible = True
    Else
        Me.lblUnitPriceEmpty.Visible = False
    End If
    
    '견적일자가 입력되었는지 체크
    If Trim(Me.txtEstimateDate.Value) = "" Then
        bCorrect = False
        Me.lblEstimateDateEmpty.Visible = True
    Else
        Me.lblEstimateDateEmpty.Visible = False
    End If
    
    If bCorrect = False Then
        Me.lblInputFieldError.Visible = True
    End If
    
    CheckEstimateInsertValidation = bCorrect
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
