Attribute VB_Name = "Mod_JoinEstimateAccepted"
Option Explicit

Sub JoinEstimateAccepted()

    Dim db As Variant
    Dim endCol, endRow As Long
    
    Application.ScreenUpdating = False
    
    ClearJoinEstimateAccepted
    
    With shtJoinEstimateAccepted
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Range("A2").Resize(endRow, endCol).Delete
    
    
        '견적 테이블과 수주 테이블을 JOIN해서 견적 테이블에 추가 필드 채움
        db = Get_DB(shtEstimateData, False, False)
        db = Join_DB(db, 2, shtAcceptedData, "관리번호", "분류1, 납기, 명세서, 계산서, 결재, 결재월, 부가세", False)
        
        '견적 테이블과 견적 메모 테이블을 JOIN 해서 메모 필드 채움
        db = Join_DB(db, 2, shtEstimateMemoData, "관리번호", "메모", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,28, 29, 30, 31, 32"
        
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = endRow
        
    End With
    
End Sub


Sub ClearJoinEstimateAccepted()

    Dim endCol, endRow As Long
    
    With shtJoinEstimateAccepted
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub
