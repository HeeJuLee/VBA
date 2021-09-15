Attribute VB_Name = "Mod_XtractDateBugEstimate"
Option Explicit

Sub ExtractDateBugEstimate()

    Dim db As Variant
    Dim i As Integer
    Dim 견적, 수주, 납품, 명세서, 계산서, 결제, 결제월 As Variant
    
    Application.ScreenUpdating = False
    
    ClearExtractDateBugEstimate
    
    With shtJoinEstimateAccepted
    
        '결제이력 추출
        db = Get_DB(shtJoinEstimateAccepted, False, False)
        For i = 1 To UBound(db)
            견적 = db(i, 12)
            수주 = db(i, 14)
            납품 = db(i, 15)
            명세서 = db(i, 27)
            계산서 = db(i, 28)
            결제 = db(i, 29)
            결제월 = db(i, 30)
            
            '견적: 37, 수주: 38, 명세서: 39, 계산서: 40, 결제: 41
            If 결제 <> "" Then
                If 견적 <> "" Then
                    If 결제 < 견적 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), 견적, , , , 결제
                    End If
                End If
                If 수주 <> "" Then
                    If 결제 < 수주 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , 수주, , , 결제
                    End If
                End If
                If 명세서 <> "" Then
                    If 결제 < 명세서 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , 명세서, , 결제
                    End If
                End If
                If 계산서 <> "" Then
                    If 결제 < 계산서 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , , 계산서, 결제
                    End If
                End If
            End If
            
            If 계산서 <> "" Then
                If 견적 <> "" Then
                    If 계산서 < 견적 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), 견적, , , 계산서
                    End If
                End If
                If 수주 <> "" Then
                    If 계산서 < 수주 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , 수주, , 계산서
                    End If
                End If
                If 명세서 <> "" Then
                    If 계산서 < 명세서 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , 명세서, 계산서
                    End If
                End If
            End If
            
            If 명세서 <> "" Then
                If 견적 <> "" Then
                    If 명세서 < 견적 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), 견적, , 명세서
                    End If
                End If
                If 수주 <> "" Then
                    If 명세서 < 수주 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , 수주, 명세서
                    End If
                End If
            End If
            
        Next
        
    End With
    
End Sub


Sub ClearExtractDateBugEstimate()

    Dim endCol, endRow As Long
    
    With shtExtractDateBugEstimate
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub

