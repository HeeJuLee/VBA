Attribute VB_Name = "Mod_JoinOrderEstimate"
Option Explicit

Sub JoinOrderEstimate()

    Dim DB As Variant
    Dim endCol, endRow As Long
    
    With shtJoinOrderEstimate
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Range("A2").Resize(endRow, endCol).Delete
    
    
        '���� ���̺�� ���� ���̺��� JOIN�ؼ� ���� ���̺� �߰� �ʵ� ä��
        DB = Get_DB(shtOrderData, False, False)
        DB = Join_DB(DB, 4, shtEstimateData, "������ȣ", "ID, �ŷ�ó, �����, ������", False)
        
        ArrayToRng .Range("A2"), DB, "1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28"
        
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = endRow
        
    End With
    
End Sub

