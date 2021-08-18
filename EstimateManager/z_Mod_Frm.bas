Attribute VB_Name = "z_Mod_Frm"
Option Explicit

'########################
' �޺��ڽ��� DB ������ ����
' Update_Cbo cboBox, DB, "1"
'########################
Sub Update_Cbo(cboBox As MSForms.ComboBox, DB As Variant, Optional DisplayCol As Long = 1, Optional SetDefault As Boolean = False)

Dim colCount As Long
Dim colWidths As String
Dim i As Long

colCount = UBound(DB, 2)

With cboBox
    .ColumnCount = colCount
    For i = 1 To colCount
        If DisplayCol = i Then colWidths = colWidths & .Width - 15 & "," Else colWidths = colWidths & "0,"
    Next
    colWidths = left(colWidths, Len(colWidths) - 1)
    .List = DB
    .ColumnWidths = colWidths
    If SetDefault = True Then .ListIndex = 0
End With

End Sub

'########################
' �޺��ڽ��� Ư�� �ʵ� ���� �����Ͽ� ���� ����
' Select_CboItm cboBox, 1, 1
'########################
Sub Select_CboItm(cboBox As MSForms.ComboBox, ID, Optional ColNo As Long = 1)

Dim i As Long

If IsNumeric(ID) Then ID = CLng(ID)

With cboBox
    For i = 0 To .ListCount - 1
        If .List(i, ColNo - 1) = ID Then .ListIndex = i
    Next
End With

End Sub

'########################
' ����Ʈ�ڽ��� DB ������ ����
' Update_List ListBox, DB, "0pt; 80pt; 50pt"
'########################
Sub Update_List(lstBox As MSForms.ListBox, DB As Variant, Widths As String)

With lstBox
    .Clear
    .ColumnWidths = Widths
    If Not IsEmpty(DB) Then
        .ColumnCount = UBound(DB, 2)
        .List = DB
    End If
End With

End Sub

'########################
' ����Ʈ�ڽ��� ������ ����� �迭�� ��ȯ
' Array = Get_ListItm listbox
'########################
Function Get_ListItm(lstBox As Control) As Variant

Dim i As Long: Dim j As Long
Dim vaArr As Variant

With lstBox
    If .ListIndex <> -1 Then
    ReDim vaArr(0 To .ColumnCount - 1)
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                For j = 0 To .ColumnCount - 1
                    vaArr(j) = .List(i, j)
                Next
                Exit For
            End If
        Next
    End If
End With

Get_ListItm = vaArr

End Function

'########################
' ����Ʈ�ڽ��� ù��°�ʵ� ID�� �����Ͽ� �ش� ID ���� ����
' Select_ListItm ListBox, ID
'########################
Function Select_ListItm(lstBox As Control, ID, Optional ColNo As Long = 1)

Dim i As Long

If IsNumeric(ID) Then ID = CLng(ID)

With lstBox
    For i = 0 To .ListCount - 1
        If .List(i, ColNo - 1) = ID Then .Selected(i) = True: Exit For
    Next
End With

End Function

'########################
' ����Ʈ�ڽ� Ȱ��ȭ
' Active_ListBox ( ListBox, Select_ListItm(ListBox, ID) )
'########################
Function Active_ListBox(lstBox As Control, Optional Index As Long = 0)

If lstBox.ListCount > 0 Then lstBox.Selected(Index) = True

End Function

'########################
' ���� ���õ� ���� ���� Ȯ��
' i = Get_ListIndex(ListBox)
'########################
Function Get_ListIndex(lstBox As Control)

Dim i As Long
With lstBox
    If .ListIndex <> -1 Then
        For i = 0 To .ListCount - 1
            If .Selected(i) Then Get_ListIndex = i: Exit For
        Next
    End If
End With

End Function

'########################
' ����Ʈ �ڽ��� ���õǾ� �ִ��� ���� Ȯ��
' boolean = isListBoxSelected(ListBox1)
'########################

Function isListBoxSelected(ListBox As MSForms.ListBox) As Boolean
 
Dim i As Long
 
For i = 0 To ListBox.ListCount - 1
If ListBox.Selected(i) Then isListBoxSelected = True: Exit Function
Next
 
isListBoxSelected = False
 
End Function
 
 '########################
' ���������� �ش� ��Ʈ�� ��ư �� �ʱ�ȭ
' Clear_Ctrls ( Userform1, "Label", "�̸�" )  ' ���������� "�̸�"�� ���� �� ���� ��� label ����
' ��Ʈ�� �̸����� ���ϵ�ī��(*,?) ��밡�� (��: txt* �� txt�� �����ϴ� ��� ��ư�� �ǹ�)
' ��Ʈ�� ���� :
' Label, Frame, TextBox, CommandButton, ComboBox, TabStrip, ListBox,
' MultiPage, CheckBox, ScrollBar, OptionButton, SpinButton, ToggleButton, Image
'########################
Sub Clear_Ctrls(frm As UserForm, CtlType As String, Optional Exclude As String)

Dim ctl As Control
Dim Excs As Variant: Dim Exc As Variant
Dim blnPass As Boolean
Dim vaType As Variant: Dim vType As Variant

If InStr(1, Exclude, ",") > 0 Then: Excs = Split(Exclude, ","): Else Excs = Array(Exclude)
If InStr(1, CtlType, ",") > 0 Then: vaType = Split(CtlType, ","): Else vaType = Array(CtlType)

For Each vType In vaType
    For Each ctl In frm.Controls
        If ctl.Name Like Trim(vType) Then
            blnPass = False
            For Each Exc In Excs
                If ctl.Name Like Trim(Exc) Then blnPass = True: Exit For
            Next
            If blnPass = False Then ctl.Value = ""
        End If
    Next
Next

End Sub

 '########################
' �������� ��Ʈ�� �� ����ִ� ��Ʈ���� �ִ��� Ȯ��(��������)
' blnCheck = IsEmpty_Ctrls ( Userform1, "Label", "�̸�" )  ' ���������� "�̸�"�� ���� �� ���� ��� label ����
' ��Ʈ�� �̸����� ���ϵ�ī��(*,?) ��밡�� (��: txt* �� txt�� �����ϴ� ��� ��ư�� �ǹ�)
' ��Ʈ�� ���� :
' Label, Frame, TextBox, CommandButton, ComboBox, TabStrip, ListBox,
' MultiPage, CheckBox, ScrollBar, OptionButton, SpinButton, ToggleButton, Image
'########################
Function IsEmpty_Ctrls(frm As UserForm, CtlType As String, Optional Exclude As String)

Dim ctl As Control
Dim vaType As Variant: Dim vType As Variant

If InStr(1, CtlType, ",") > 0 Then: vaType = Split(CtlType, ","): Else vaType = Array(CtlType)

For Each vType In vaType
    For Each ctl In frm.Controls
        If ctl.Name Like Trim(vType) And ctl.Name <> Exclude Then
            If ctl.Value = "" Then IsEmpty_Ctrls = True: Exit Function
        End If
    Next
Next

IsEmpty_Ctrls = False

End Function
