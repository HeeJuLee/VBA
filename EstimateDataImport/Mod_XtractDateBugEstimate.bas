Attribute VB_Name = "Mod_XtractDateBugEstimate"
Option Explicit

Sub ExtractDateBugEstimate()

    Dim db As Variant
    Dim i As Integer
    Dim ����, ����, ��ǰ, ����, ��꼭, ����, ������ As Variant
    
    Application.ScreenUpdating = False
    
    ClearExtractDateBugEstimate
    
    With shtJoinEstimateAccepted
    
        '�����̷� ����
        db = Get_DB(shtJoinEstimateAccepted, False, False)
        For i = 1 To UBound(db)
            ���� = db(i, 12)
            ���� = db(i, 14)
            ��ǰ = db(i, 15)
            ���� = db(i, 27)
            ��꼭 = db(i, 28)
            ���� = db(i, 29)
            ������ = db(i, 30)
            
            '����: 37, ����: 38, ����: 39, ��꼭: 40, ����: 41
            If ���� <> "" Then
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), ����, , , , ����
                    End If
                End If
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , ����, , , ����
                    End If
                End If
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , ����, , ����
                    End If
                End If
                If ��꼭 <> "" Then
                    If ���� < ��꼭 Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , , ��꼭, ����
                    End If
                End If
            End If
            
            If ��꼭 <> "" Then
                If ���� <> "" Then
                    If ��꼭 < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), ����, , , ��꼭
                    End If
                End If
                If ���� <> "" Then
                    If ��꼭 < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , ����, , ��꼭
                    End If
                End If
                If ���� <> "" Then
                    If ��꼭 < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , , ����, ��꼭
                    End If
                End If
            End If
            
            If ���� <> "" Then
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), ����, , ����
                    End If
                End If
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugEstimate, db(i, 1), db(i, 2), db(i, 6), , ����, ����
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

