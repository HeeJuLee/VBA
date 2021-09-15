Attribute VB_Name = "Mod_XtractDataBugOrder"
Option Explicit

Sub ExtractDateBugOrder()

    Dim db As Variant
    Dim i As Integer
    Dim ����, �԰�, ����, ��꼭, ����, ������ As Variant
    
    Application.ScreenUpdating = False
    
    ClearExtractDateBugOrder
    
    With shtJoinOrderEstimate
    
        '�����̷� ����
        db = Get_DB(shtJoinOrderEstimate, False, False)
        For i = 1 To UBound(db)
            ���� = db(i, 16)
            �԰� = db(i, 18)
            ���� = db(i, 20)
            ��꼭 = db(i, 21)
            ���� = db(i, 22)
            ������ = db(i, 23)
            
            If ���� <> "" Then
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), ����, , , , ����
                    End If
                End If
                If �԰� <> "" Then
                    If ���� < �԰� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , �԰�, , , ����
                    End If
                End If
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , ����, , ����
                    End If
                End If
                If ��꼭 <> "" Then
                    If ���� < ��꼭 Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , , ��꼭, ����
                    End If
                End If
            End If
            
            If ��꼭 <> "" Then
                If ���� <> "" Then
                    If ��꼭 < ���� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), ����, , , ��꼭
                    End If
                End If
                If �԰� <> "" Then
                    If ��꼭 < �԰� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , �԰�, , ��꼭
                    End If
                End If
                If ���� <> "" Then
                    If ��꼭 < ���� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , , ����, ��꼭
                    End If
                End If
            End If
            
            If ���� <> "" Then
                If ���� <> "" Then
                    If ���� < ���� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), ����, , ����
                    End If
                End If
                If �԰� <> "" Then
                    If ���� < �԰� Then
                        Insert_Record shtExtractDateBugOrder, db(i, 1), db(i, 2), db(i, 7), , �԰�, ����
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


