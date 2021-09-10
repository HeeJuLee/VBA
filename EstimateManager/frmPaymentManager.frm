VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPaymentManager 
   Caption         =   "결제 이력 관리"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "frmPaymentManager.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmPaymentManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private acceptedPrice As String

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
    If paymentFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = paymentFormX
        Me.top = paymentFormY
    End If
    
    'currentEstimateId로 견적데이터 읽어오기 (확인용)
    estimate = Get_Record_Array(shtEstimate, currentEstimateId)
    If IsEmpty(estimate) Then
        MsgBox "currentEstimateId에 해당하는 견적 데이터가 없습니다.", vbInformation, "작업 확인"
        End
    End If

    Me.txtEstimateName.value = estimate(6)
    Me.txtManagementID.value = estimate(2)
    acceptedPrice = estimate(21)
    
    InitializeLswPaymentList    '결제 이력
    
    ClearPaymentInput
End Sub

Sub InitializeLswPaymentList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '견적ID에 해당하는 결제이력을 읽어옴
    db = Get_DB(shtPayment)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
     '리스트뷰 값 설정
    With Me.lswPaymentList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .HideSelection = True
        .FullRowSelect = True
        .MultiSelect = True
        .LabelEdit = lvwManual
        .CheckBoxes = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "결제일", 70, lvwColumnCenter
        .ColumnHeaders.Add , , "결제금액", 70, lvwColumnRight
        .ColumnHeaders.Add , , "메모", 140
        .ColumnHeaders.Add , , "등록일자", 0
        
        '.ColumnHeaders(1).Position = 1
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If IsNumeric(db(i, 5)) Then
                    '비용 합계 구함
                    totalCost = totalCost + CLng(db(i, 5))
                End If
                
                Set li = .ListItems.Add(, , db(i, 1))
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , Format(db(i, 5), "#,##0")
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                
                li.Selected = False
            Next
            
            Me.txtPaid.value = Format(totalCost, "#,##0")
            If IsNumeric(acceptedPrice) Then
                Me.txtRemaining.value = Format(acceptedPrice - totalCost, "#,##0")
            End If
        End If
    End With
End Sub

Sub InsertPayment()
    
    If Me.txtPayDate.value = "" Then MsgBox "결제일을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    If Me.txtPayAmount.value = "" Then MsgBox "결제금액을 입력하세요.", vbInformation, "작업 확인": Exit Sub

    '분할결제이력에 저장
    Insert_Record shtPayment, CLng(currentEstimateId), Me.txtManagementID.value, Me.txtPayDate.value, Me.txtPayAmount.value, Me.txtPayMemo.value, Date
    
    '합계 계산
    Me.txtPaid.value = Format(GetPaymentTotalCost, "#,##0")
    If IsNumeric(acceptedPrice) Then
        Me.txtRemaining.value = Format(acceptedPrice - Me.txtPaid.value, "#,##0")
    End If
    
    '입금액/미입금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "입금액", Me.txtPaid.value
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "미입금액", Me.txtRemaining.value
    
    '입금액/미입금액을 frmEstimateUpdate 폼에도 업데이트
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtPaymentDate.value = Me.txtPayDate.value
        frmEstimateUpdate.txtPaid.value = Me.txtPaid.value
        frmEstimateUpdate.txtRemaining.value = Me.txtRemaining.value
    End If
    
    InitializeLswPaymentList
    
    '등록한 아이템 선택
    Me.txtPayID.value = Get_LastID(shtPayment)
    SelectItemLswPayment Me.txtPayID.value
    
End Sub


Sub UpdatePayment()
    Dim cost As Variant

    If Me.txtPayID.value = "" Then MsgBox "수정할 항목을 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    If Me.txtPayDate.value = "" Then MsgBox "결제일을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    If Me.txtPayAmount.value = "" Then MsgBox "결제금액을 입력하세요.", vbInformation, "작업 확인": Exit Sub
    
    '기존 분할결제이력에 업데이트
    Update_Record shtPayment, Me.txtPayID.value, currentEstimateId, Me.txtManagementID.value, Me.txtPayDate.value, Me.txtPayAmount.value, Me.txtPayMemo.value, Date
    
    '합계 계산
    Me.txtPaid.value = Format(GetPaymentTotalCost, "#,##0")
    If IsNumeric(acceptedPrice) Then
        Me.txtRemaining.value = Format(acceptedPrice - Me.txtPaid.value, "#,##0")
    End If
    
    '입금액/미입금액을 견적테이블에 저장
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "입금액", Me.txtPaid.value
    Update_Record_Column shtEstimate, CLng(currentEstimateId), "미입금액", Me.txtRemaining.value
    
    '입금액/미입금액을 frmEstimateUpdate 폼에도 업데이트
    If isFormLoaded("frmEstimateUpdate") Then
        frmEstimateUpdate.txtPaymentDate.value = Me.txtPayDate.value
        frmEstimateUpdate.txtPaid.value = Me.txtPaid.value
        frmEstimateUpdate.txtRemaining.value = Me.txtRemaining.value
    End If
    
    InitializeLswPaymentList
    SelectItemLswPayment Me.txtPayID.value
    
End Sub


Sub DeletePayment()
    Dim db As Variant
    Dim YN As VbMsgBoxResult
    Dim count As Long
    Dim li As ListItem

    count = 0
    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then count = count + 1
    Next
    If count = 0 Then MsgBox "삭제할 항목을 선택하세요.", vbInformation, "작업 확인": Exit Sub
    
    YN = MsgBox("선택한 " & count & "개 항목을 삭제할까요?", vbYesNo + vbQuestion, "작업 확인")
    If YN = vbNo Then Exit Sub

    For Each li In Me.lswPaymentList.ListItems
        If li.Selected = True Then
            '결제이력 테이블에서 삭제
            Delete_Record shtPayment, li.Text
        End If
    Next
    
    If count > 0 Then
        '합계 계산
        Me.txtPaid.value = Format(GetPaymentTotalCost, "#,##0")
        If IsNumeric(acceptedPrice) Then
            Me.txtRemaining.value = Format(acceptedPrice - Me.txtPaid.value, "#,##0")
        End If
        
        '입금액/미입금액을 견적테이블에 저장
        Update_Record_Column shtEstimate, CLng(currentEstimateId), "입금액", Me.txtPaid.value
        Update_Record_Column shtEstimate, CLng(currentEstimateId), "미입금액", Me.txtRemaining.value
        
        '입금액/미입금액을 frmEstimateUpdate 폼에도 업데이트
        If isFormLoaded("frmEstimateUpdate") Then
            frmEstimateUpdate.txtPaymentDate.value = Me.txtPayDate.value
            frmEstimateUpdate.txtPaid.value = Me.txtPaid.value
            frmEstimateUpdate.txtRemaining.value = Me.txtRemaining.value
        End If
    End If
        
    Me.txtPayID.value = ""
    InitializeLswPaymentList
    ClearPaymentInput
    
End Sub
Function GetPaymentTotalCost()
    Dim i As Long
    Dim totalCost As Long
    Dim db As Variant
    
    '견적ID에 해당하는 결제이력을 읽어옴
    db = Get_DB(shtPayment)
    db = Filtered_DB(db, currentEstimateId, 2, True)
    
    'DB에 값이 있을 경우
    totalCost = 0
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db)
            If IsNumeric(db(i, 5)) Then
                '비용 합계 구함
                totalCost = totalCost + CLng(db(i, 5))
            End If
        Next
    End If
        
    GetPaymentTotalCost = totalCost
End Function

Sub SelectItemLswPayment(selectedID As Variant)
    Dim i As Long
    
    With Me.lswPaymentList
        If Not IsMissing(selectedID) Then
            For i = 1 To .ListItems.count
                If selectedID = .ListItems(i).Text Then
                    .selectedItem = .ListItems(i)
                    .selectedItem.EnsureVisible
                    .SetFocus
                Else
                    .ListItems(i).Selected = False
                End If
            Next
        End If
    End With
End Sub

Sub ClearPaymentInput()
    Me.txtPayID.value = ""
    Me.txtPayDate.value = ""
    Me.txtPayAmount.value = ""
    Me.txtPayMemo.value = ""
End Sub

Private Sub btnPaymentClear_Click()
    ClearPaymentInput
End Sub

Private Sub btnPaymentDelete_Click()
    DeletePayment
End Sub

Private Sub btnPaymentInsert_Click()
    InsertPayment
End Sub

Private Sub btnPaymentUpdate_Click()
    UpdatePayment
End Sub


Private Sub btnPaymentClose_Click()
    Unload Me
End Sub

Private Sub lswPaymentList_Click()
    With Me.lswPaymentList
        If Not .selectedItem Is Nothing Then
            Me.txtPayID.value = .selectedItem.Text
            Me.txtPayDate.value = .selectedItem.ListSubItems(1)
            Me.txtPayAmount.value = .selectedItem.ListSubItems(2)
            Me.txtPayMemo.value = .selectedItem.ListSubItems(3)
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

Private Sub btnPaymentClear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then
        Me.btnPaymentClose.SetFocus
    End If
End Sub

Private Sub imgPayDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GetCalendarDate Me.txtPayDate
End Sub


Private Sub txtPayAmount_AfterUpdate()
    Me.txtPayAmount.value = Trim(Me.txtPayAmount.value)
    
    If Not IsNumeric(Me.txtPayAmount.value) Then
        MsgBox "숫자를 입력하세요.", vbInformation, "작업 확인"
        Exit Sub
    End If
    
    Me.txtPayAmount.value = Format(Me.txtPayAmount.value, "#,##0")
End Sub

Private Sub txtPayDate_AfterUpdate()
    Me.txtPayDate.value = Trim(Me.txtPayDate.value)
End Sub

Private Sub txtPayDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtPayMemo_AfterUpdate()
    Me.txtPayMemo.value = Trim(Me.txtPayMemo.value)
End Sub


Private Sub UserForm_Layout()
    paymentFormX = Me.Left
    paymentFormY = Me.top
End Sub

