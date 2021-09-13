VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductionCopy 
   Caption         =   "예상실행항목으로 가져오기"
   ClientHeight    =   10725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "frmProductionCopy.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmProductionCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim selectedEstimateId As Variant

Private Sub UserForm_Initialize()
    Dim contr As Control
    Dim estimate As Variant
    
    '폼 위치 수정
    If productionCopyFormX <> 0 Then
        Me.StartUpPosition = 0
        Me.Left = productionCopyFormX
        Me.top = productionCopyFormY
    End If
    
    '텍스트박스 라벨 위치 조정
    For Each contr In Me.Controls
        If contr.Name Like "Label*" Then
            contr.top = contr.top + 2
        End If
    Next
    
    InitializeLswEstimateList
    InitializeLswOrderList
    
End Sub

Sub InitializeLswEstimateList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '리스트뷰 값 설정
    With Me.lswEstimateList
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
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "관리번호", 76
        .ColumnHeaders.Add , , "거래처", 60
        .ColumnHeaders.Add , , "담당자", 50
        .ColumnHeaders.Add , , "견적명", 230
        .ColumnHeaders.Add , , "수주일자", 60, lvwColumnCenter
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswEstimateList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '키워드로 검색한 견적 목록을 가져옴
    If Me.txtKeyword.value <> "" Then
        db = Get_DB(shtEstimate)
        db = Filtered_DB(db, Me.txtKeyword.value)
        If Not IsEmpty(db) Then
            db = Filtered_DB(db, Me.txtKeyword2.value)
        End If
        
        With Me.lswEstimateList
            .ListItems.Clear
            If Not IsEmpty(db) Then
                For i = 1 To UBound(db)
                    Set li = .ListItems.Add(, , db(i, 1))
                    li.ListSubItems.Add , , db(i, 2)
                    li.ListSubItems.Add , , db(i, 4)
                    li.ListSubItems.Add , , db(i, 5)
                    li.ListSubItems.Add , , db(i, 6)
                    li.ListSubItems.Add , , db(i, 14)
                    
                    li.Selected = False
                Next
            End If
        End With
        
        selectedEstimateId = ""
    End If
End Sub

Sub InitializeLswOrderList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    With Me.ImageList1.ListImages
        .Add , , Me.imgListImage.Picture
    End With
    
     '리스트뷰 값 설정
    With Me.lswOrderList
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
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "분류", 34
        .ColumnHeaders.Add , , "거래처", 68
        .ColumnHeaders.Add , , "품목", 100
        .ColumnHeaders.Add , , "재질", 50
        .ColumnHeaders.Add , , "규격", 50
        .ColumnHeaders.Add , , "수량", 44, lvwColumnRight
        .ColumnHeaders.Add , , "단위", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "단가", 60, lvwColumnRight
        .ColumnHeaders.Add , , "금액", 60, lvwColumnRight
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswOrderList(estimateId)
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '견적ID에 해당하는 발주리스트를 읽어옴
    db = Get_DB(shtOrder)
    db = Filtered_DB(db, estimateId, 28, True)
    
     '리스트뷰 값 설정
    With Me.lswOrderList
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                If db(i, 4) <> "수주" Then
                    Set li = .ListItems.Add(, , db(i, 1))
                    li.ListSubItems.Add , , db(i, 4)
                    li.ListSubItems.Add , , db(i, 6)
                    li.ListSubItems.Add , , db(i, 7)
                    li.ListSubItems.Add , , db(i, 8)
                    li.ListSubItems.Add , , db(i, 9)
                    li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                    li.ListSubItems.Add , , db(i, 11)
                    li.ListSubItems.Add , , Format(db(i, 12), "#,##0")
                    li.ListSubItems.Add , , Format(db(i, 13), "#,##0")
                    li.Selected = False
                End If
            Next
            
        End If
    End With
End Sub

Sub InitializeLswProductionList()
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
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
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_견적", 0
        .ColumnHeaders.Add , , "관리번호", 0
        .ColumnHeaders.Add , , "분류", 34
        .ColumnHeaders.Add , , "거래처", 68
        .ColumnHeaders.Add , , "품목", 100
        .ColumnHeaders.Add , , "재질", 50
        .ColumnHeaders.Add , , "규격", 50
        .ColumnHeaders.Add , , "수량", 44, lvwColumnRight
        .ColumnHeaders.Add , , "단위", 30, lvwColumnCenter
        .ColumnHeaders.Add , , "단가", 60, lvwColumnRight
        .ColumnHeaders.Add , , "금액", 60, lvwColumnRight
        .ColumnHeaders.Add , , "메모", 0
        
        .ListItems.Clear
    End With
End Sub

Sub SetLswProductionList(estimateId)
    Dim db As Variant
    Dim i, j, totalCost As Long
    Dim li As ListItem
        
    '견적ID에 해당하는 예상비용항목을 읽어옴
    db = Get_DB(shtProduction)
    db = Filtered_DB(db, estimateId, 2, True)
    
     '리스트뷰 값 설정
    With Me.lswProductionList
    
        .ListItems.Clear
        If Not IsEmpty(db) Then
            For i = 1 To UBound(db)
                Set li = .ListItems.Add(, , db(i, 1))
                li.ListSubItems.Add , , db(i, 13)
                li.ListSubItems.Add , , db(i, 2)
                li.ListSubItems.Add , , db(i, 3)
                li.ListSubItems.Add , , db(i, 4)
                li.ListSubItems.Add , , db(i, 5)
                li.ListSubItems.Add , , db(i, 6)
                li.ListSubItems.Add , , db(i, 7)
                li.ListSubItems.Add , , Format(db(i, 8), "#,##0")
                li.ListSubItems.Add , , db(i, 9)
                li.ListSubItems.Add , , Format(db(i, 10), "#,##0")
                li.ListSubItems.Add , , Format(db(i, 11), "#,##0")
                li.ListSubItems.Add , , db(i, 12)
                
                li.Selected = False
            Next
            
        End If
    End With
End Sub

Private Sub lswEstimateList_Click()
    With Me.lswEstimateList
        If Not .selectedItem Is Nothing Then
            selectedEstimateId = .selectedItem.Text
            SetLswOrderList (selectedEstimateId)
        End If
    End With
End Sub

Private Sub lswEstimateList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lswEstimateList
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
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

Private Sub btnEstimateSearch_Click()
    SetLswEstimateList
End Sub

Private Sub btnProductionClose_Click()
    Unload Me
End Sub

Private Sub btnProductionCopyAll_Click()
    ProductionCopy "all"
End Sub

Private Sub btnProductionCopy_Click()
    ProductionCopy ""
End Sub

Sub ProductionCopy(all)
    Dim count As Long
    Dim yn As Variant
    Dim li As ListItem
    
    If selectedEstimateId = "" Then
        MsgBox "복사할 견적을 선택하세요.", vbInformation, "작업 확인"
        Exit Sub
    End If
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If all = "all" Then
            count = count + 1
        Else
            If li.Selected = True Then count = count + 1
        End If
    Next
    If count = 0 Then MsgBox "복사할 항목을 선택하세요.": Exit Sub
        
    yn = MsgBox(count & "개 항목을 복사할까요?", vbYesNo + vbQuestion, "작업 확인")
    If yn = vbNo Then Exit Sub
    
    count = 0
    For Each li In Me.lswOrderList.ListItems
        If li.Selected = True Or all = "all" Then
            Insert_Record shtProduction, currentEstimateId, currentManagementId, li.SubItems(2), li.SubItems(3), li.SubItems(4), li.SubItems(5), li.SubItems(6), li.SubItems(7), li.SubItems(8), li.SubItems(9), , li.SubItems(1), Date
            count = count + 1
        End If
    Next
    
    If isFormLoaded("frmProductionManager") Then
        frmProductionManager.RefreshProductionTotalCost
    End If
    
    MsgBox count & "개 항목을 복사하였습니다.", vbInformation, "작업 확인"
    
End Sub

Private Sub txtKeyword_AfterUpdate()
    Me.txtKeyword.value = Trim(Me.txtKeyword.value)
    SetLswEstimateList
End Sub

Private Sub txtKeyword2_AfterUpdate()
    Me.txtKeyword2.value = Trim(Me.txtKeyword2.value)
    SetLswEstimateList
End Sub

Private Sub UserForm_Layout()
    productionCopyFormX = Me.Left
    productionCopyFormY = Me.top
End Sub

