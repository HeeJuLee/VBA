Attribute VB_Name = "Test"
Option Explicit

Sub test()
'
'Update_Record_Column shtTest, 15, "담당자", "이희주"
'Update_Record_Column shtTest, 15, "연락처", "123-331-333"
'Update_Record_Column shtEstimate, 35, "수정일자", Date
'Update_Record_Column shtEstimate, 35, "실행가", 3000000

Stop
    Dim arr As Variant
    
    arr = Get_Record_Array(shtEstimate, 35)
  
    ArrayToRng shtTest.Range("B2"), arr
    
End Sub


Sub 견적_테스트_데이터_생성()

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
    
    
    '기존 조회내용 지우기
    shtTest.Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    '조회할SQL을 만들어서 String변수에 넣는다.
    'strSQL = "SELECT * FROM [견적$] WHERE [등록일자] > '2021-08-01' AND [등록일자] <= '2021-08-31' "
    'strSQL = "SELECT * FROM [견적$] WHERE [견적명] LIKE '%부품%' or [관리번호] LIKE '%부품%' or [자재번호] LIKE '%부품%'"
    
    Dim start_date As Date
    start_date = #8/21/2021#
    
    start_date = Format(Date, "yyyy-mm-dd hh:mm:ss")
    
    'strSQL = "SELECT [ID], [견적명], [등록일자] FROM [견적$] WHERE [등록일자] = '" & start_date & "' and [견적명] like '%부품%' "
    strSQL = "SELECT [ID], [견적명], [등록일자] FROM [견적$] WHERE cvdate([등록일자]) = '" & start_date & "' "
  

           
    'Excel을 Database로 사용
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & ActiveWorkbook.Name & ";" & "Extended Properties=Excel 12.0;"
    
    rs.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    

    If rs.EOF Then
        MsgBox "조회조건에 해당하는 자료가 없습니다."
    Else
        '타이틀을 표시한다.
        For i = 1 To rs.fields.count
          Cells(1, i).Value = rs.fields(i - 1).Name
        Next
        
        
        With ActiveSheet
            '조회한 결과집합(rs)을 "출력"Sheet의 A2지점을 꼭지점으로 해서 출력한다.
           .Range("A2").CopyFromRecordset rs
        End With
    End If
    
   
    rs.Close
    Set rs = Nothing
    
End Sub
