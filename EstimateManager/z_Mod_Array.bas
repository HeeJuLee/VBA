Attribute VB_Name = "z_Mod_Array"
Sub ArrayToRng(startRng As Range, arr As Variant, Optional ColumnNo As String = "")

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ArrayToRng �Լ�
'�� �迭�� ���� ���� ��ȯ�մϴ�.
'�� �μ� ����
'_____________startRng      : �迭�� ��ȯ�� ���� ����(��) �Դϴ�.
'_____________Arr               : ��ȯ�� �迭�Դϴ�.
'_____________ColumnNo   : [�����μ�] �迭�� Ư�� ���� �����Ͽ� ������ ��ȯ�մϴ�. ������ ���� ��ȯ�� ��� �� ��ȣ�� ��ǥ�� �����Ͽ� �Է��մϴ�.
'                                               ������ ������ �Է��ϸ� ���� �ǳʶݴϴ�.
'�� ��� ����
'Dim v As Variant
'ReDim v(0 to 1)
''v(0) = "a" : v(1) = "b"
'ArrayToRng Sheet1.Range("A1"), v
'�� ���� ���� ��ɹ�
'Extract_Column �Լ�
'##############################################################

On Error GoTo SingleDimension:

Dim Cols As Variant: Dim col As Variant
Dim x As Long: x = 1
If ColumnNo = "" Then
    startRng.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1) = arr
Else
    Cols = Split(ColumnNo, ",")
    For Each col In Cols
        If Trim(col) <> "" Then
            startRng.Cells(1, x).Resize(UBound(arr, 1) - LBound(arr, 1) + 1) = Extract_Column(arr, CLng(Trim(col)))
        End If
        x = x + 1
    Next
End If
Exit Sub

SingleDimension:
Dim tempArr As Variant: Dim i As Long
ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), 1 To 1)
For i = LBound(arr, 1) To UBound(arr, 1)
    tempArr(i, 1) = arr(i)
Next
startRng.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, 1) = tempArr

End Sub

Sub SequenceToRng(startRng As Range, count As Long, Optional StartNo As Double = 1, Optional Increment As Double = 1, Optional ToRight As Boolean = False)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� SequenceToRng �Լ�
'�� ������ ������ ��ȯ�մϴ�.
'�� �μ� ����
'_____________startRng      : �迭�� ��ȯ�� ���� ����(��) �Դϴ�.
'_____________Count          : �迭�� ��ȯ�� ������ �����Դϴ�.
'_____________StartNo        : [�����μ�] ������ ���� ��ȣ�Դϴ�. �⺻���� 1 �Դϴ�.
'_____________Increment    : [�����μ�]������ ���� �Ǵ� �����ϴ� ���̰��Դϴ�. �⺻���� 1 �Դϴ�.
'_____________ToRight        : [�����μ�] True�� ��� ������ ������ �������� ��ȯ�մϴ�. �⺻���� False(=�Ʒ�����)�Դϴ�.
'�� ��� ����
'SequenceToRng Range("A1")
'##############################################################

Dim arr As Variant: Dim v As Double: v = StartNo - Increment

If ToRight = False Then ReDim arr(1 To count, 1 To 1) Else ReDim arr(1 To 1, 1 To count)

If ToRight = False Then
    For i = 1 To count
        v = v + Increment
        arr(i, 1) = v
    Next
Else
    For i = 1 To count
        v = v + Increment
        arr(1, i) = v
    Next
End If

If ToRight = False Then startRng.Cells(1, 1).Resize(count) = arr Else startRng.Cells(1, 1).Resize(1, count) = arr

End Sub

Sub ValueToRng(startRng As Range, count As Long, value, Optional ToRight As Boolean = False)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ValueToRng �Լ�
'�� ������ ���� ������ �ѷ��ݴϴ�.
'�� �μ� ����
'_____________startRng      : ���� �ѷ��� ���� �� �Դϴ�.
'_____________Count          : �ѷ��� ���� �����Դϴ�.
'_____________Value           : �ѷ��� �� �Դϴ�.
'_____________ToRight        : [�����μ�] True�� ��� ���� ������ �������� �ѷ��ݴϴ�. �⺻���� False(=�Ʒ�����)�Դϴ�.
'�� ��� ����
'ValueToRng Range("A1"), 10, "A"  '<- A1:A10 ������ "A"�� ����մϴ�.
'##############################################################

If ToRight = False Then startRng.Cells(1, 1).Resize(count) = value Else startRng.Cells(1, 1).Resize(1, count) = value

End Sub

Sub RunningSumRng(startRng As Range, count As Long, _
                                    Optional Offset_Add As Long = -1, Optional Offset_Deduct As Long = 0, _
                                    Optional blnReverse As Boolean = False)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� RunningSumRng �Լ�
'�� ���ۼ��� �������� ������ġ/����ġ�� �����Ͽ� ���踦 ����մϴ�.
'�� �μ� ����
'_____________startRng               : ���踦 ����� ���� ���Դϴ�.
'_____________Count                   : ���踦 ����� ������ �����Դϴ�.
'_____________Offset_Add          : ���� ���� �Էµ� ���� ��� ��ġ�Դϴ�. (������ ���ʹ���, ����� �����ʹ����Դϴ�. �⺻���� -1, ��ĭ ���ʿ� �ִ� ���� ���մϴ�.)
'_____________Offset_Deduct     : [�����μ�] �� ���� �Էµ� ���� �����ġ�Դϴ�. �⺻���� �����Դϴ�.
'_____________blnReverse            : [�����μ�] True�� ��� ���ۼ��� �� �Ʒ��� ���� �����Ͽ� ���� �ö󰡸� ���踦 ����մϴ�. �⺻���� False �Դϴ�.
'�� ��� ����
'RunningSumRng Range("C1"), 10,   '<- C1:C10 ������ B1:B10�� �����Ͽ� ���踦 ����մϴ�.
'##############################################################

Dim T As Double
Dim vArr As Variant
Dim fR As Single: Dim fS As Long: Dim fE As Long

If blnReverse = False Then fR = 1 Else fR = -1
If count < 1 Then count = 1

ReDim vArr(1 To count, 1 To 1)
    If fR = 1 Then fS = 1: fE = count Else fS = count: fE = 1
    
    If Offset_Deduct <> 0 Then
        For i = 1 To count
            T = T + startRng.Offset((i - 1) * fR, Offset_Add).value - startRng.Offset((i - 1) * fR, Offset_Deduct).value
            vArr(i, 1) = T
        Next
    Else
        For i = 1 To count
            T = T + startRng.Offset((i - 1) * fR, Offset_Add).value
            vArr(i, 1) = T
        Next
    End If

If fR = 1 Then
    startRng.Resize(count) = vArr
Else
    fE = UBound(vArr, 1)
    fS = LBound(vArr, 1)
    For i = fS To (fE - fS) \ 2 + fS
        T = vArr(fE, 1)
        vArr(fE, 1) = vArr(i, 1)
        vArr(i, 1) = T
        fE = fE - 1
    Next
    startRng.Offset(-count + 1).Resize(count) = vArr
End If

End Sub

Function IsUniqueArray(arr As Variant, Optional colNo As String = "") As Boolean

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� IsUniqueArray �Լ�
'�� �迭���� ������ ���� �ϳ��� �����θ� �ԷµǾ� �ִ��� Ȯ���մϴ�.
'�� �μ� ����
'_____________Arr               : �迭�Դϴ�.
'_____________ColNo          : [�����μ�] �ϳ��� ���� �ԷµǾ� �ִ��� Ȯ���� �� ��ȣ�Դϴ�. ��ǥ�� �����Ͽ� �������� ���� ������ �� �ֽ��ϴ�. AND �������� �������θ� ��ȸ�մϴ�. �⺻���� �迭�� ���� ����ȣ�Դϴ�.
'�� ��� ����
'Debug.Print IsUniqueArray (DB, 2)  '<- �迭�� �ι�° ���� ���� �ϳ��θ� �̷�����ִ��� Ȯ���մϴ�.
'�� ���� �����Լ�
'ArrayDimension �Լ�
'##############################################################

Dim D As Long: Dim i As Long
Dim c As Variant
Dim vCols As Variant: Dim vCol As Variant
Dim sTemp As String

D = ArrayDimension(arr)

If colNo = "" Then If D > 1 Then colNo = LBound(arr, 2)
vCols = Split(colNo, ",")

If D = 1 Then
    c = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> c Then IsUniqueArray = False: Exit Function
    Next
Else
    For Each vCol In vCols
        sTemp = sTemp & arr(LBound(arr, 1), CLng(Trim(vCol)))
    Next
    c = sTemp
    For i = LBound(arr, 1) To UBound(arr, 1)
        sTemp = ""
        For Each vCol In vCols
            sTemp = sTemp & arr(i, CLng(Trim(vCol)))
        Next
        If c <> sTemp Then IsUniqueArray = False: Exit Function
    Next
End If

IsUniqueArray = True

End Function
Function IsDistinctArray(arr As Variant, Optional colNo As String = "") As Boolean

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� IsDistinctArray �Լ�
'�� �迭���� ������ ���� ������ ��� ������ ���鸸 �ԷµǾ� �ִ��� Ȯ���մϴ�.
'�� �μ� ����
'_____________Arr               : �迭�Դϴ�.
'_____________ColNo          : [�����μ�] ������ �������� ���θ� ��ȸ�� �� ��ȣ�Դϴ�. ��ǥ�� �����Ͽ� �������� ���� ������ �� �ֽ��ϴ�. AND �������� �������θ� ��ȸ�մϴ�. �⺻���� �迭�� ���� ����ȣ�Դϴ�.
'�� ��� ����
'Debug.Print IsDistinctArray(DB, 2)  '<- �迭�� �ι�° ���� ������ �������� �˻��մϴ�.
'�� ���� �����Լ�
'ArrayDimension �Լ�
'##############################################################

Dim Dict As Dictionary
Dim vCols As Variant: Dim vCol As Variant
Dim sTemp As String: Dim i As Long

Set Dict = New Dictionary

If colNo = "" Then If ArrayDimension(arr) > 1 Then colNo = LBound(arr, 2)

vCols = Split(colNo, ",")

On Error GoTo Duplicate:

If ArrayDimension(arr) = 1 Then
    For i = LBound(arr) To UBound(arr)
        Dict.Add arr(i), 0
    Next
Else
    For i = LBound(arr, 1) To UBound(arr, 1)
        sTemp = ""
        For Each vCol In vCols
            sTemp = sTemp & arr(i, CLng(Trim(vCol)))
        Next
        Dict.Add sTemp, 0
    Next
End If

IsDistinctArray = True

Exit Function

Duplicate:
    IsDistinctArray = False

End Function

Sub ClearContentsBelow(startRng As Range, Optional colNo, Optional BaseCol As Long = 0)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ClearContentsBelow �Լ�
'�� ���ؼ����� Ư�������� �Ʒ��� �Էµ� �����͸� ��� �ʱ�ȭ�մϴ�. (���� �ʱ�ȭ�Ǹ� ������ �״�� �����˴ϴ�.)
'�� �μ� ����
'_____________startRng      : ���ؼ��Դϴ�.
'_____________ColNo          : [�����μ�] ���ؼ��κ��� ������ ����ȣ(�Ǵ� ���ĺ�)�Դϴ�. �⺻���� ���ؼ��κ��� ���ӵ� ������ ���� �������� ����ȣ�� ��ȯ�մϴ�.
'_____________BaseCol       : [�����μ�] ���ؼ� �Ʒ��� ���ӵ� �����͸� ������ ���ؿ���ȣ �Դϴ�. �⺻���� ���ؼ��� ����ȣ�Դϴ�.
'�� ��� ����
'ClearContentsBelow Range("A5"), "F"   '<- A5~F������ �Ʒ��� �Էµ� �����͸� �ʱ�ȭ�մϴ�.
'##############################################################

Dim WS As Worksheet: Dim lastRow As Long: Set WS = startRng.Parent
If IsMissing(colNo) Then colNo = WS.Cells(startRng.row, WS.Columns.count).End(xlToLeft).Column
If Not IsNumeric(colNo) Then colNo = Range(colNo & 1).Column
If BaseCol = 0 Then BaseCol = startRng.Column Else BaseCol = startRng.Column + BaseCol - 1
lastRow = WS.Cells(WS.Rows.count, BaseCol).End(xlUp).row
If lastRow < startRng.row Then Exit Sub
WS.Range(startRng, WS.Cells(lastRow, colNo)).ClearContents

End Sub

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ArrayDimension �Լ�
'�� �迭�� �������� ��ȯ�մϴ�.
'�� �μ� ����
'_____________vaArray     : ������ ������ �迭�� �Է��մϴ�.
'###############################################################
Function ArrayDimension(vaArray As Variant) As Integer
 
Dim i As Integer: Dim x As Integer
 
On Error Resume Next
 
Do
    i = i + 1
    x = UBound(vaArray, i)
Loop Until Err.Number <> 0
 
Err.Clear
 
ArrayDimension = i - 1
 
End Function

