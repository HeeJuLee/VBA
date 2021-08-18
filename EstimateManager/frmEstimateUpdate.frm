VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstimateUpdate 
   Caption         =   "견적 정보 수정"
   ClientHeight    =   7515
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

Private Sub cboCustomer_Change()
    '콤보박스에서 거래처를 변경하면 해당 거래처의 담당자로 담당자 콤보박스를 세팅
    InitializeCboManager
End Sub


Private Sub imgAcceptedDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtAcceptedDate
End Sub

Private Sub imgBidDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtBidDate
End Sub

Private Sub imgDeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtDeliveryDate
End Sub

Private Sub imgEstimateDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtEstimateDate
End Sub

Private Sub imgInsuranceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtInsuranceDate
End Sub

Private Sub imgPaymentDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPaymentDate
End Sub

Private Sub imgSpecificationDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtSpecificationDate
End Sub

Private Sub imgTaxInvoiceDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtTaxInvoiceDate
End Sub

Private Sub txtEstimateName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub


'버튼 마우스오버 처리
'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnEstimateUpdate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEstimateUpdate
End Sub

Private Sub btnEstimateUpdate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEstimateUpdate
End Sub

Private Sub btnEstimateUpdate_Enter()
OnHover_Css Me.btnEstimateUpdate
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
Dim btnList As String: btnList = "btnEstimateUpdate, btnEstimateClose" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
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


Private Sub UserForm_Initialize()
    Dim cRow As Long
    Dim estimate As Variant
    Dim manager As Variant
    Dim customer As Variant
    
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
    
    Me.txtAmount.Value = estimate(7)    '수량
    Me.txtUnitPrice.Value = estimate(9)     '견적단가
    Me.txtEstimatePrice.Value = estimate(10)    '견적금액
    
    Me.txtEstimateDate.Value = estimate(11)    '견적일자
    Me.txtBidDate.Value = estimate(12)    '입찰일자
    Me.txtAcceptedDate.Value = estimate(13)    '수주일자
    Me.txtDeliveryDate.Value = estimate(14)    '납품일자
    Me.txtInsuranceDate.Value = estimate(15)    '증권일자
    
    Me.txtProductionTotalCost.Value = estimate(16)    '예상실행가
    Me.txtBidPrice.Value = estimate(17)    '입찰가
    Me.txtBidMargin.Value = estimate(18)    '차액
    Me.txtBidMarginRate.Value = estimate(19)    '마진율
    Me.txtAcceptedPrice.Value = estimate(20)    '수주금액
    Me.txtAcceptedMargin.Value = estimate(21)    '수주차액
    
    Me.txtSpecificationDate.Value = estimate(22)    '거래명세서
    Me.txtTaxInvoiceDate.Value = estimate(23)    '세금계산서
    Me.txtPaymentDate.Value = estimate(24)    '결제일
    Me.txtPaymentMonth.Value = estimate(25)    '예상결제월
    Me.txtVAT.Value = estimate(26)    '부가세
    Me.txtProjection.Value = estimate(27)    '입금예상액
    Me.txtPaid.Value = estimate(28)    '입금액
    Me.txtUnpaid.Value = estimate(29)    '미입금액
    
    Me.txtInsertDate.Value = estimate(30)    '등록일자
    Me.txtUpdateDate.Value = estimate(31)    '수정일자
    
    '변경 전 관리번호
    orgEstimateID = Me.txtEstimateID
    
End Sub


Sub UpdateEstimate()
    Dim DB As Variant
    Dim blnUnique As Boolean
    
    '입력 데이터 체크
'    If InputValidationCheck = False Then
'        Exit Sub
'    End If

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
        Me.txtPaymentMonth.Value, Me.txtVAT.Value, _
        Me.txtProjection.Value, Me.txtPaid.Value, _
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



