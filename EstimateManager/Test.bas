Attribute VB_Name = "Test"
Option Explicit

Sub test()
'
'Update_Record_Column shtTest, 15, "�����", "������"
'Update_Record_Column shtTest, 15, "����ó", "123-331-333"
'Update_Record_Column shtEstimate, 35, "��������", Date
'Update_Record_Column shtEstimate, 35, "���డ", 3000000

Stop
    Dim arr As Variant
    
    arr = Get_Record_Array(shtEstimate, 35)
  
    ArrayToRng shtTest.Range("B2"), arr
    
End Sub


Sub ����_�׽�Ʈ_������_����()

    Dim i, j As Long
    Dim startRow, endRow, setCount As Long
    
   j = 202
    For i = 1 To 60
        Range("A30").Resize(172, 31).Copy
        
        Range("A" & j).Resize(172, 31).Insert
        
        j = j + 172
    Next

End Sub

Sub TestADO()
    
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim strConn As String
Dim i As Integer
    
    
    '���� ��ȸ���� �����
    shtTest.Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    '��ȸ��SQL�� ���� String������ �ִ´�.
    'strSQL = "SELECT * FROM [����$] WHERE [�������] > '2021-08-01' AND [�������] <= '2021-08-31' "
    'strSQL = "SELECT * FROM [����$] WHERE [������] LIKE '%��ǰ%' or [������ȣ] LIKE '%��ǰ%' or [�����ȣ] LIKE '%��ǰ%'"
    
    Dim start_date As Date
    start_date = #8/21/2021#
    
    start_date = Format(Date, "yyyy-mm-dd hh:mm:ss")
    
    'strSQL = "SELECT [ID], [������], [�������] FROM [����$] WHERE [�������] = '" & start_date & "' and [������] like '%��ǰ%' "
    strSQL = "SELECT [ID], [������], [�������] FROM [����$] WHERE cvdate([�������]) = '" & start_date & "' "
  

           
    'Excel�� Database�� ���
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & ActiveWorkbook.Name & ";" & "Extended Properties=Excel 12.0;"
    
    rs.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    

    If rs.EOF Then
        MsgBox "��ȸ���ǿ� �ش��ϴ� �ڷᰡ �����ϴ�."
    Else
        'Ÿ��Ʋ�� ǥ���Ѵ�.
        For i = 1 To rs.fields.count
          Cells(1, i).Value = rs.fields(i - 1).Name
        Next
        
        
        With ActiveSheet
            '��ȸ�� �������(rs)�� "���"Sheet�� A2������ ���������� �ؼ� ����Ѵ�.
           .Range("A2").CopyFromRecordset rs
        End With
    End If
    
   
    rs.Close
    Set rs = Nothing
    
End Sub
