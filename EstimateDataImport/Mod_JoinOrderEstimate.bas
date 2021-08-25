Attribute VB_Name = "Mod_JoinOrderEstimate"
Option Explicit

Sub JoinOrderEstimate()

    Dim db As Variant
    Dim endCol, endRow As Long
    
    With shtJoinOrderEstimate
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Range("A2").Resize(endRow, endCol).Delete
    
    
        '발주 테이블과 견적 테이블을 JOIN해서 발주 테이블에 추가 필드 채움
        db = Get_DB(shtOrderData, False, False)
        db = Join_DB(db, 4, shtEstimateData, "관리번호", "ID", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25"
        
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = endRow
        
    End With
    
End Sub

