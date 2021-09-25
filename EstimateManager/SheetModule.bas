Attribute VB_Name = "SheetModule"
Option Explicit

Sub UpdateShtOrderField(orderId, fieldName, fieldValue)
    Dim findRow, colNo As Long
    
    findRow = isExistInSheet(shtOrderAdmin.Range("B6"), orderId)
    If findRow > 0 Then
        Select Case fieldName
            Case "분류1"
                colNo = 7
            Case "거래처"
                colNo = 8
            Case "품목"
                colNo = 9
            Case "재질"
                colNo = 10
            Case "규격"
                colNo = 11
            Case "수량"
                colNo = 12
            Case "단위"
                colNo = 13
            Case "단가"
                colNo = 14
            Case "금액"
                colNo = 15
            Case "중량"
                colNo = 16
            Case "수주"
                colNo = 17
            Case "발주"
                colNo = 18
            Case "납기"
                colNo = 19
            Case "입고"
                colNo = 20
            Case "납품"
                colNo = 21
            Case "명세서"
                colNo = 22
            Case "계산서"
                colNo = 23
            Case "결제"
                colNo = 24
            Case "결제월"
                colNo = 25
            Case "결제수단"
                colNo = 26
            Case "부가세"
                colNo = 27
            Case "등록일자"
                colNo = 28
            Case "수정일자"
                colNo = 29
        End Select
        
        If colNo > 0 Then
            shtOrderAdmin.Cells(findRow, colNo).value = fieldValue
        End If
    End If
End Sub


