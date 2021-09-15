Attribute VB_Name = "Mod_JoinOrderEstimate"
Option Explicit

Sub JoinOrderEstimate()

    Dim db As Variant
    Dim endCol, endRow As Long
    Dim estimateId As Variant
    Dim managementId As Variant
    Dim spec, tax, paid, month, acceptedPrice, payMethod, vat, regDate, memo As Variant
    Dim paidPrice As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ClearJoinOrderEstimate
    
    With shtJoinOrderEstimate
    
        '발주 테이블과 견적 테이블을 JOIN해서 발주 테이블에 추가 필드 채움
        db = Get_DB(shtOrderData, False, False)
        db = Join_DB(db, 5, shtEstimateData, "관리번호", "ID, 납품", False)
        
        '발주 테이블과 관리 메모 테이블을 JOIN해서 메모 필드 채움
        db = Join_DB(db, 2, shtManageMemoData, "ID_관리", "메모", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30"
     
    End With
    
    '결제이력 추출
    db = Get_DB(shtJoinOrderEstimate, False, False)
    For i = 1 To UBound(db)
        If db(i, 4) = "수주" Then
            estimateId = db(i, 28)
            managementId = db(i, 5)
            If InStr(db(i, 9), "%") > 0 Or InStr(db(i, 9), "금") > 0 Or InStr(db(i, 9), "액") > 0 Then
                memo = db(i, 9)
            Else
                memo = ""
            End If
            acceptedPrice = db(i, 13)
            spec = db(i, 20)
            tax = db(i, 21)
            paid = db(i, 22)
            month = db(i, 23)
            payMethod = db(i, 24)
            vat = db(i, 25)
            regDate = db(i, 26)
            
            If IsNumeric(acceptedPrice) Then paidPrice = CLng(acceptedPrice) Else paidPrice = 0
            
            
            If acceptedPrice <> "" And paid <> "" Then
                '금액이 있고 결제가 있는 경우
                Insert_Record shtPaymentData, estimateId, managementId, spec, tax, paid, month, paidPrice, "", payMethod, memo, vat, regDate, ""
            ElseIf acceptedPrice <> "" And paid = "" And month <> "" Then
                '금액이 있고 결제가 없고 결제월이 있는 경우
                Insert_Record shtPaymentData, estimateId, managementId, spec, tax, paid, month, "", paidPrice, payMethod, memo, vat, regDate, ""
            End If
        End If
    Next
    
End Sub

Sub ClearJoinOrderEstimate()

    Dim endCol, endRow As Long
    
    With shtJoinOrderEstimate
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).ClearContents
    End With
    
    With shtPaymentData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).ClearContents
    End With
End Sub
