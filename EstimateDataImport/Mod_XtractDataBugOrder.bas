Attribute VB_Name = "Mod_XtractDataBugOrder"
Option Explicit

Sub ExtractDateBugOrder()

    Dim db As Variant
    Dim i As Integer
    Dim 발주, 입고, 명세서, 계산서, 결제, 결제월 As Variant
    
    Application.ScreenUpdating = False
    
    ClearExtractDateBugOrder
    
    With shtJoinOrderEstimate
    
        '결제이력 추출
        db = Get_DB(shtJoinOrderEstimate, False, False)
        For i = 1 To UBound(db)
            발주 = db(i, 16)
            입고 = db(i, 18)
            명세서 = db(i, 20)
            계산서 = db(i, 21)
            결제 = db(i, 22)
            결제월 = db(i, 23)
            
            If 결제 <> "" Then
                If 발주 <> "" Then
                    If 결제 < 발주 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), 발주, , , , 결제
                    End If
                End If
                If 입고 <> "" Then
                    If 결제 < 입고 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , 입고, , , 결제
                    End If
                End If
                If 명세서 <> "" Then
                    If 결제 < 명세서 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , 명세서, , 결제
                    End If
                End If
                If 계산서 <> "" Then
                    If 결제 < 계산서 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , , 계산서, 결제
                    End If
                End If
            End If
            
            If 계산서 <> "" Then
                If 발주 <> "" Then
                    If 계산서 < 발주 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), 발주, , , 계산서
                    End If
                End If
                If 입고 <> "" Then
                    If 계산서 < 입고 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , 입고, , 계산서
                    End If
                End If
                If 명세서 <> "" Then
                    If 계산서 < 명세서 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , 명세서, 계산서
                    End If
                End If
            End If
            
            If 명세서 <> "" Then
                If 발주 <> "" Then
                    If 명세서 < 발주 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), 발주, , 명세서
                    End If
                End If
                If 입고 <> "" Then
                    If 명세서 < 입고 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , 입고, 명세서
                    End If
                End If
            End If
            
        Next
        
    End With
    
End Sub


Sub ClearExtractDateBugOrder()

    Dim endCol, endRow As Long
    
    With shtExtractDateBugOrder
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub


