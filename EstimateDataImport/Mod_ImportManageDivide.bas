Attribute VB_Name = "Mod_ImportManageDivide"
Option Explicit

Sub DivideManage()

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    ClearManageDivide
    
    With shtManageData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
    
        For i = 2 To endRow
            man.ID = .Cells(i, 1)
            man.수입지출 = .Cells(i, 2)
            man.분류1 = .Cells(i, 3)
            man.분류2 = .Cells(i, 4)
            man.관리번호 = .Cells(i, 5)
            man.거래처 = .Cells(i, 6)
            man.품목 = .Cells(i, 7)
            man.재질 = .Cells(i, 8)
            man.규격 = .Cells(i, 9)
            man.단가 = .Cells(i, 10)
            man.금액 = .Cells(i, 11)
            man.단위 = .Cells(i, 12)
            man.중량 = .Cells(i, 13)
            man.수량 = .Cells(i, 14)
            man.수주 = .Cells(i, 15)
            man.납기 = .Cells(i, 16)
            man.발주 = .Cells(i, 17)
            man.입고 = .Cells(i, 18)
            man.납품 = .Cells(i, 19)
            man.명세서 = .Cells(i, 20)
            man.계산서 = .Cells(i, 21)
            man.결재 = .Cells(i, 22)
            man.결재월 = .Cells(i, 23)
            man.부가세 = .Cells(i, 24)
            man.등록일자 = .Cells(i, 25)
            
            If man.수입지출 = "수입" And man.분류2 = "수주" Then
                '수입이면서 수주이면 수주 테이블에 등록
                Insert_Record shtAcceptedData, man.ID, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자
                
            ElseIf man.수입지출 = "지출" And Len(man.관리번호) >= 10 Then
                '지출이면서 관리번호가 있으면 발주 테이블에 등록
                Insert_Record shtOrderData, man.ID, man.분류2, man.관리번호, man.거래처, man.품목, man.재질, man.규격, man.수량, man.단위, man.단가, man.금액, man.중량, _
                              man.발주, man.납기, man.입고, man.명세서, man.계산서, man.결재, man.결재월, man.분류1, man.부가세, man.등록일자
                              
            Else
                '그 외 남는 것은 운영비 테이블에 등록
                Insert_Record shtOperatingData, man.ID, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, man.금액, man.명세서, man.결재, man.부가세, man.등록일자
            End If
            
        Next
    End With
End Sub

Sub ClearManageDivide()

    Dim endCol, endRow As Long
    
    With shtAcceptedData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtOrderData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtOperatingData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub
