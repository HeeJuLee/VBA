VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductionManager 
   Caption         =   "예상실행항목 관리"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   OleObjectBlob   =   "frmProductionManager.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmProductionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim estimate As Variant
    
    If currentEstimateId = "" Then
        MsgBox "currentEstimateId 오류: 선택한 견적이 없습니다.", vbInformation, "작업 확인"
        End
    End If
    
    '텍스트박스 라벨 컨트롤 색상 조정
    For Each contr In Me.Controls
        If contr.Name Like "lbl*" Then
            If contr.Name Like "lbl2*" Then
                'contr.BackColor = RGB(48, 84, 150)
                'contr.ForeColor = RGB(255, 255, 255)
            ElseIf contr.Name Like "lbl3*" Then
                contr.BackColor = RGB(221, 235, 247)
            Else
                contr.BackColor = RGB(242, 242, 242)
            End If
        End If
    Next
    
    '폼 위치 수정
    If productionFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = productionFormX
        Me.top = productionFormY
    End If
    
    'currentEstimateId로 견적데이터 읽어오기 (확인용)
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If IsEmpty(estimate) Then
        MsgBox "currentEstimateId에 해당하는 견적 데이터가 없습니다.", vbInformation, "작업 확인"
        End
    End If

    Me.txtEstimateName.value = estimate(6)
    Me.txtManagementID.value = estimate(2)
    Me.txtCustomer.value = estimate(4)
    Me.txtManager.value = estimate(5)
    
    InitializeCboCategory           '분류
    InitializeLswProductionList    '예상실행항목 목록
    InitializeCboProductonUnit  '예상실행항목 단위
    InitializeLswOrderCustomerAutoComplete   '발주거래처 자동완성
    
    ClearProductionInput
End Sub

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '리스트뷰 값 설정
    With Me.lswProductionList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwAutomatic
        .CheckBoxes = False
        .SmallIcons = Me.ImageList1
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "품목", 130
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "분류", 34
        .ColumnHeaders.Add , , "거래처", 70
        .ColumnHeaders.Add , , "재질", 60
        .ColumnHeaders.Add , , "규격", 80
        .ColumnHeaders.Add , , "수량", 44, lvwColumnRight
        .ColumnHeaders.Add , , "단위", 44, lvwColumnCenter
        .ColumnHeaders.Add , , "단가", 70, lvwColumnRight
        .ColumnHeaders.Add , , "금액", 70, lvwColumnRight
        .ColumnHeaders.Add , , "메모", 92
        .ColumnHeaders.Add , , "등록일자", 0
        
        .ColumnHeaders(1).Position = 6
    
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
                li.ListSubItems.Add , , db(i, 13)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , db(i, 8)
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                
                li.Selected = False
            Next
            
            Me.txtProductionTotalCost.value = Format(totalCost, "#,##0")
        End If
    End With
End Sub

Sub InitializeLswOrderCustomerAutoComplete()
    With Me.lswOrderCustomerAutoComplete
        .View = lvwList
        .LabelEdit = lvwManual
        .Height = 126
        .Visible = False
    End With
End Sub

Sub InitializeCboCategory()
    Dim db As Variant
    db = Get_DB(shtOrderCategory, True)

    Update_Cbo Me.cboCategory, db
End Sub

Sub InitializeCboProductonUnit()
    Dim db As Variant
    db = Get_DB(shtUnit, True)

    Update_Cbo Me.cboProductionUnit, db
End Sub


Sub InsertProduction()
    
    If Me.txtProductionItem.value = "" Then MsgBox "품목을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "금액을 입력하세요.", vbInformation, "작업 확인": Exit Sub

    '예상실행항목에 저장
    Insert_Record shtProduction, CLng(currentEstimateId), Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Me.cboCategory.value, Date
    
    '예상실행가를 포함한 곳에 업데이트
    RefreshProductionTotalCost
    
    '등록한 아이템 선택
    Me.txtProductionID.value = Get_LastID(shtProduction)
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub


Sub UpdateProduction()
    Dim cost As Variant

    If Me.txtProductionID.value = "" Then MsgBox "수정할 항목을 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    If Me.txtProductionItem.value = "" Then MsgBox "품목을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    If Me.txtProductionCost.value = "" Then MsgBox "금액을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    
    '기존 예상실행항목에 업데이트
    Update_Record shtProduction, Me.txtProductionID.value, currentEstimateId, Me.txtManagementID.value, Me.txtProductionCustomer.value, Me.txtProductionItem.value, _
            Me.txtProductionMaterial.value, Me.txtProductionSize.value, _
            Me.txtProductionAmount.value, Me.cboProductionUnit.value, Me.txtProductionUnitPrice.value, Me.txtProductionCost.value, Me.txtProductionMemo.value, Me.cboCategory.value, Date
    
    '예상실행가를 포함한 곳에 업데이트
    RefreshProductionTotalCost
    
    SelectItemLswProduction Me.txtProductionID.value
    
End Sub

Sub RefreshProductionTotalCost()
    '예상실행항목 합계 계산
    Me.txtProductionTotalCost.value = Format(GetProductionTotalCost, "#,##0")
    
    '예상실행가를 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "실행가(예상)", CLng(Me.txtProductionTotalCost.value)
    
    '예상실행가를 frmEstimateUpdate 폼 값도 업데이트
    If isFormLoaded("frmEstimateUpdate") Then
        'frmEstimateUpdate.txtProductionTotalCost = Me.txtProductionTotalCost.value
        frmEstimateUpdate.UpdateProductionTotalCost Me.txtProductionTotalCost.value
'        frmEstimateUpdate.CalculateEstimateUpdateCost
'        frmEstimateUpdate.UpdateShtEstimateField currentEstimateId, "예상실행가", Me.txtProductionTotalCost.value
    End If
    
    '예상실행항목 리스트박스 새로고침
    InitializeLswProductionList
End Sub

Sub DeleteProduction()
    Dim db As Variant
    Dim YN As VbMsgBoxResult
    Dim count As Long
    Dim li As ListItem

    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "삭제할 항목을 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    YN = MsgBox("선택한 " & count & "개 항목을 삭제할까요?", vbYesNo + vbQuestion, "작업 확인")
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            '예상실행항목 테이블에서 삭제
            Delete_Record shtProduction, li.SubItems(1)
        End If
    Next
    
    '예상실행가를 포함한 곳에 업데이트
    RefreshProductionTotalCost
        
    Me.txtProductionID.value = ""
    ClearProductionInput
    
End Sub

Sub ProductionToOrder()
    Dim li As ListItem
    Dim count As Long
    Dim managementId, category, customer, Item, material, size, amount, unit, unitPrice, cost, memo As Variant
    Dim YN As VbMsgBoxResult
    Dim estimate As Variant
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "발주할 항목을 선택하세요.": Exit Sub
    
    '수주 확정이 아닌 견적의 경우는 발주를 할 수 없음
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If estimate(38) = "" Then
        MsgBox "수주 확정을 진행해야 발주할 수 있습니다.", vbInformation, "작업 확인"
        Exit Sub
    End If
    
    YN = MsgBox("선택한 " & count & "개 항목을 발주할까요?", vbYesNo + vbQuestion, "작업 확인")
    If YN = vbNo Then Exit Sub
    
    count = 0
    For Each li In Me.lswProductionList.ListItems
        If li.Selected = True Then
            Item = li.Text
            managementId = li.SubItems(3)
            category = li.SubItems(4)
            customer = li.SubItems(5)
            material = li.SubItems(6)
            size = li.SubItems(7)
            amount = li.SubItems(8)
            unit = li.SubItems(9)
            unitPrice = li.SubItems(10)
            cost = li.SubItems(11)
            memo = li.SubItems(12)
            
            '선택한 예상실행항목을 발주 테이블에 등록
            Insert_Record shtOrder, _
                , , category, managementId, customer, Item, material, size, amount, unit, unitPrice, cost, , _
                , , , , , _
                , , , , _
                , , _
                Date, , currentEstimateId, memo, False
                
            count = count + 1
        End If
    Next
    
    'frmEstimateUpdate 폼의 발주목록을 업데이트
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.InitializeLswOrderList
        frmEstimateUpdate.CalculateEstimateUpdateCost
    End If
    
    MsgBox "총 " & count & "개 항목을 발주하였습니다.", vbInformation, "작업 확인"
    
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
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
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

Sub SelectItemLswProduction(selectedID As Variant)
    Dim i As Long
    
    With Me.lswProductionList
        If Not IsMissing(selectedID) Then
            For i = 1 To .ListItems.count
                If selectedID = .ListItems(i).SubItems(1) Then
                    .SetFocus
                    .selectedItem = .ListItems(i)
                    .selectedItem.EnsureVisible
                Else
                    .ListItems(i).Selected = False
                End If
            Next
        End If
    End With
End Sub

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

Private Sub btnProductionCopy_Click()
    If isFormLoaded("frmProductionCopy") Then
        Unload frmProductionCopy
    End If
    frmProductionCopy.Show (False)
End Sub

Private Sub btnProductionClose_Click()
    Unload Me
End Sub

Private Sub lswProductionList_Click()
    With Me.lswProductionList
        If Not .selectedItem Is Nothing Then
            Me.txtProductionItem.value = .selectedItem.Text
            Me.txtProductionID.value = .selectedItem.ListSubItems(1)
            Me.cboCategory.value = .selectedItem.ListSubItems(4)
            Me.txtProductionCustomer.value = .selectedItem.ListSubItems(5)
            Me.txtProductionMaterial.value = .selectedItem.ListSubItems(6)
            Me.txtProductionSize.value = .selectedItem.ListSubItems(7)
            Me.txtProductionAmount.value = .selectedItem.ListSubItems(8)
            Me.cboProductionUnit.value = .selectedItem.ListSubItems(9)
            Me.txtProductionUnitPrice.value = .selectedItem.ListSubItems(10)
            Me.txtProductionCost.value = .selectedItem.ListSubItems(11)
            Me.txtProductionMemo.value = .selectedItem.ListSubItems(12)
        End If
    End With
End Sub

Private Sub lswProductionList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswProductionList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub


Private Sub btnProductionClear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnProductionClose.SetFocus
    End If
End Sub

Private Sub cboCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtProductionCustomer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        '엔터키 - 다음 입력칸으로 이동
        Me.lswOrderCustomerAutoComplete.Visible = False
        Me.txtProductionItem.SetFocus
    ElseIf KeyCode = 9 Then
        '탭키 자동완성이 하나이면 다음으로 이동
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count = 1 Then
                Me.lswOrderCustomerAutoComplete.Visible = False
                Me.txtProductionItem.SetFocus
                KeyCode = 0
            Else
                If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
            End If
        End With
    ElseIf KeyCode = 40 Then
        '아래화살키 - 자동완성 결과가 있는 경우에는 포커스를 자동완성 리스트로 이동
        With Me.lswOrderCustomerAutoComplete
            If .ListItems.count > 0 And .Visible = True Then
                .selectedItem = .ListItems(1)
                .SetFocus
            End If
        End With
    ElseIf KeyCode = 27 Then
        'ESC키 닫기
        Unload Me
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
    '거래처에 값을 넣어주고 포커스는 품목으로 이동
    With Me.lswOrderCustomerAutoComplete
        If Not .selectedItem Is Nothing Then
            Me.txtProductionCustomer.value = .selectedItem.Text
            .Visible = False
            Me.txtProductionItem.SetFocus
        End If
    End With
End Sub

Private Sub lswOrderCustomerAutoComplete_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    '거래처에 값을 넣어주고 포커스는 품목으로 이동
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


Private Sub txtProductionItem_Enter()
    If Me.lswOrderCustomerAutoComplete.Visible = True Then
        With Me.lswOrderCustomerAutoComplete
            Me.txtProductionCustomer.value = .selectedItem.Text
            .Visible = False
        End With
    End If
End Sub


Private Sub txtProductionCustomer_AfterUpdate()
    Me.txtProductionCustomer.value = Trim(Me.txtProductionCustomer.value)
End Sub


Private Sub cboCategory_AfterUpdate()
    Me.cboCategory.value = Trim(Me.cboCategory.value)
End Sub


Private Sub txtProductionAmount_AfterUpdate()
    If Me.txtProductionAmount.value = "" Then
        Me.txtProductionCost.value = Me.txtProductionUnitPrice.value
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionAmount.value) Then
        MsgBox "숫자를 입력하세요.", vbInformation, "작업 확인"
        Me.txtProductionAmount.value = ""
        Me.txtProductionCost.value = Me.txtProductionUnitPrice.value
        Exit Sub
    End If
        
    Me.txtProductionAmount.value = Format(Me.txtProductionAmount.value, "#,##0")
        
    '금액 = 수량 * 단가
    If IsNumeric(Me.txtProductionUnitPrice.value) Then
        Me.txtProductionCost.value = Format(CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value), "#,##0")
    End If
End Sub


Private Sub txtProductionItem_AfterUpdate()
    Me.txtProductionItem.value = Trim(Me.txtProductionItem.value)
End Sub

Private Sub txtProductionMaterial_AfterUpdate()
    Me.txtProductionMaterial.value = Trim(Me.txtProductionMaterial.value)
End Sub


Private Sub txtProductionMemo_AfterUpdate()
    Me.txtProductionMemo.value = Trim(Me.txtProductionMemo.value)
End Sub


Private Sub txtProductionSize_AfterUpdate()
    Me.txtProductionSize.value = Trim(Me.txtProductionSize.value)
End Sub


Private Sub txtProductionUnitPrice_AfterUpdate()
    If Me.txtProductionUnitPrice.value = "" Then
        Me.txtProductionCost.value = ""
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtProductionUnitPrice.value) Then
        MsgBox "숫자를 입력하세요.", vbInformation, "작업 확인"
        Me.txtProductionUnitPrice.value = ""
        Me.txtProductionCost.value = ""
        Exit Sub
    End If
    
    Me.txtProductionUnitPrice.value = Format(Me.txtProductionUnitPrice.value, "#,##0")
    
    If IsNumeric(Me.txtProductionUnitPrice.value) Then
        If Me.txtProductionAmount.value = "" Then
            Me.txtProductionCost.value = Format(Me.txtProductionCost.value, "#,##0")
        Else
            If IsNumeric(Me.txtProductionAmount.value) Then
                '금액 = 수량 * 단가
                Me.txtProductionCost.value = Format(CLng(Me.txtProductionAmount.value) * CLng(Me.txtProductionUnitPrice.value), "#,##0")
            End If
        End If
    End If
    
End Sub

Private Sub UserForm_Layout()
    productionFormX = Me.Left
    productionFormY = Me.top
End Sub

