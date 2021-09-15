Attribute VB_Name = "Mod_JoinEstimateAccepted"
Option Explicit

Sub JoinEstimateAccepted()

    Dim db As Variant
    Dim endCol, endRow As Long
    Dim i, pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    Dim insDate As String
    
    
    Application.ScreenUpdating = False
    
    ClearJoinEstimateAccepted
    
    With shtJoinEstimateAccepted
   
    
        '견적 테이블과 수주 테이블을 JOIN해서 견적 테이블에 추가 필드 채움
        db = Get_DB(shtEstimateData, False, False)
        db = Join_DB(db, 2, shtAcceptedData, "관리번호", "분류1, 납기, 명세서, 계산서, 결재, 결재월, 부가세, ID_관리", False)
        
        '견적 테이블과 견적 메모 테이블을 JOIN 해서 견적메모 필드 채움
        db = Join_DB(db, 2, shtEstimateMemoData, "관리번호", "메모", False)
        
        '견적 테이블과 수주 메모 테이블을 JOIN해서 수주메모 필드 채움
        db = Join_DB(db, 2, shtAcceptedMemoData, "관리번호", "메모", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,28, 29, 30, 31, 32, 33, 34"
        
        
        
        '등록일자를 관리번호에서 추출해서 업데이트
        '        db = Get_DB(shtJoinEstimateAccepted, False, False)
        '        For i = 1 To UBound(db)
        '            managementId = db(i, 2)
        '            regDate = db(i, 23)
        '            insDate = ""
        '
        '            If insDate = "" Then
        '                pos = InStr(managementId, "-")
        '                If pos > 0 Then
        '                    M = Mid(managementId, pos - 4, 2)
        '                    D = Mid(managementId, pos - 2, 2)
        '                    If IsNumeric(M) And IsNumeric(D) Then
        '                        Y = year(regDate)
        '                        insDate = DateSerial(Y, M, D)
        '                    End If
        '                End If
        '            End If
        '
        '            If insDate <> "" Then
        '                regDate = insDate
        '            End If
        '
        '            Update_Record_Column shtJoinEstimateAccepted, db(i, 1), "등록일자", regDate
        '        Next
        
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
