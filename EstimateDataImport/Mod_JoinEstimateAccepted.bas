Attribute VB_Name = "Mod_JoinEstimateAccepted"
Option Explicit

Sub JoinEstimateAccepted()

    Dim db As Variant
    Dim endCol, endRow As Long
    
    With shtJoinEstimateAccepted
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Range("A2").Resize(endRow, endCol).Delete
    
    
        '���� ���̺�� ���� ���̺��� JOIN�ؼ� ���� ���̺� �߰� �ʵ� ä��
        db = Get_DB(shtEstimateData, False, False)
        db = Join_DB(db, 2, shtAcceptedData, "������ȣ", "�з�1, ����, ��꼭, ����, �����, �ΰ���", False)
        
        ArrayToRng .Range("A2"), db, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,28, 29, 30"
        
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = endRow
        
    End With
    
End Sub
