Attribute VB_Name = "Mod_JoinEstimateAccepted"
Option Explicit

Sub JoinEstimateAccepted()

    Dim db As Variant
    Dim endCol, endRow As Long
    Dim i, pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    Dim insDate As String
    Dim estimateId As Variant
    Dim managementId As Variant
    Dim spec, tax, paid, month, acceptedPrice, vat, regDate As Variant
    Dim paidPrice As Long
    
    Application.ScreenUpdating = False
    
    ClearJoinEstimateAccepted
    
    With shtJoinEstimateAccepted
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Range("A2").Resize(endRow, endCol).Delete
    
    
        '견적 테이블과 수주 테이블을 JOIN해서 견적 테이블에 추가 필드 채움
        db = Get_DB(shtEstimateData, False, False)
        db = Join_DB(db, 2, shtAcceptedData, "관리번호", "분류1, 납기, 명세서, 계산서, 결재, 결재월, 부가세, ID_관리", False)
        
        '견적 테이블과 견적 메모 테이블을 JOIN 해서 견적메모 필드 채움
        db = Join_DB(db, 2, shtEstimateMemoData, "관리번호", "메모", False)
        
        '견적 테이블과 수주 메모 테이블을 JOIN해서 수주메모 필드 채움
        db = Join_DB(db, 2, shtAcceptedMemoData, "관리번호", "메모", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,28, 29, 30, 31, 32, 33, 34"
        
        '결제이력 추출
        db = Get_DB(shtJoinEstimateAccepted, False, False)
        For i = 1 To UBound(db)
            estimateId = db(i, 1)
            managementId = db(i, 2)
            acceptedPrice = db(i, 21)
            spec = db(i, 27)
            tax = db(i, 28)
            paid = db(i, 29)
            month = db(i, 30)
            vat = db(i, 31)
            regDate = db(i, 23)
            
            If IsNumeric(acceptedPrice) Then paidPrice = CLng(acceptedPrice) Else paidPrice = 0
            
            '수주금액이 있고 결제가 있는 경우 - 입금액 업데이트
            If acceptedPrice <> "" And paid <> "" Then
                '입금액 업데이트
                Update_Record_Column shtJoinEstimateAccepted, estimateId, "입금액", paidPrice
                
                '결제이력에 등록
                Insert_Record shtPaymentData, estimateId, managementId, spec, tax, paid, month, paidPrice, "", "", vat, regDate, ""
            ElseIf acceptedPrice <> "" And paid = "" And month <> "" Then
                '수주금액이 있고 결제가 없고 결제월이 있는 경우 - 미입금액 업데이트
                Update_Record_Column shtJoinEstimateAccepted, estimateId, "미입금액", paidPrice
                
                '결제이력에 등록
                Insert_Record shtPaymentData, estimateId, managementId, spec, tax, paid, month, "", "", "", "", regDate, ""
            End If
            
        Next
        
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
    
    With shtPaymentData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub
