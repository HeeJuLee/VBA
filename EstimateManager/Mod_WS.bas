Attribute VB_Name = "Mod_WS"
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

Sub GetCalendarDate(textBox As MSForms.textBox)
    Dim vDate As Date
    Dim orgValue As Variant
    
    orgValue = textBox.Value
    
    vDate = frmCalendar.GetDate
    
    ' X나 ESC를 눌러서 나온 경우, 날짜 포맷이 '오전 10:00' 이런식으로 넘어옴. 오류 체크
    If InStr(vDate, "오전") <> 0 Or InStr(vDate, "오후") <> 0 Then
        'X를 누른 경우에 미리 입력되어 있던 값이 있으면 유지
        If orgValue = "" Then
            textBox.Value = ""
        End If
    Else
        textBox.Value = vDate
    End If
End Sub

