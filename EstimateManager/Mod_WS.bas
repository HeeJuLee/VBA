Attribute VB_Name = "Mod_WS"
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

Sub GetCalendarDate(textBox As MSForms.textBox)
    Dim vDate As Date
    Dim orgValue As Variant
    
    orgValue = textBox.Value
    
    vDate = frmCalendar.GetDate
    
    ' X�� ESC�� ������ ���� ���, ��¥ ������ '���� 10:00' �̷������� �Ѿ��. ���� üũ
    If InStr(vDate, "����") <> 0 Or InStr(vDate, "����") <> 0 Then
        'X�� ���� ��쿡 �̸� �ԷµǾ� �ִ� ���� ������ ����
        If orgValue = "" Then
            textBox.Value = ""
        End If
    Else
        textBox.Value = vDate
    End If
End Sub

