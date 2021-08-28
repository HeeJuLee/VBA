Attribute VB_Name = "z_Mod_DB"
Option Explicit
Option Compare Text
'########################
' Ư�� ��ũ��Ʈ���� ������ �߰��ؾ� �� �ִ� ID��ȣ ���� (��Ʈ DB ���� ù��° �Ӹ���)
' i = Get_MaxID(Sheet1)
'########################
Function Get_MaxID(WS As Worksheet) As Long
With WS
    Get_MaxID = .Cells(1, .Columns.Count).End(xlToLeft).Value
    .Cells(1, .Columns.Count).End(xlToLeft).Value = .Cells(1, .Columns.Count).End(xlToLeft).Value + 1
End With
End Function
'########################
' hjlee 2021.08.22 �߰�
' Ư�� ��ũ��Ʈ�� ���� ID��ȣ ���� (��Ʈ DB ���� ù��° �Ӹ���)
' i = Get_CurrentID(Sheet1)
'########################
Function Get_CurrentID(WS As Worksheet) As Long
With WS
    Get_CurrentID = .Cells(1, .Columns.Count).End(xlToLeft).Value
End With
End Function
'########################
' ��ũ��Ʈ�� ���ο� �����͸� �߰��ؾ� �� ����ȣ ��ȯ
' i = Get_InsertRow(Sheet1)
'########################
Function Get_InsertRow(WS As Worksheet) As Long
With WS:    Get_InsertRow = .Cells(.Rows.Count, 1).End(xlUp).row + 1: End With
End Function
'########################
' ��Ʈ�� �� ���� ��ȯ (�̹� �������Ͽ����� ���)
' i  = Get_ColumnCnt(Sheet1)
'########################
Function Get_ColumnCnt(WS As Worksheet, Optional Offset As Long = -1) As Long
With WS:    Get_ColumnCnt = .Cells(1, .Columns.Count).End(xlToLeft).Column + Offset: End With
End Function
'########################
' ��Ʈ���� Ư�� ID �� �� ��ȣ ��ȯ (-> �ش� �� ��ȣ ������ ������Ʈ)
' i = get_UpdateRow(Sheet1, ID)
'########################
Function get_UpdateRow(WS As Worksheet, ID)
Dim i As Long
Dim cRow As Long
With WS
    cRow = Get_InsertRow(WS) - 1
    For i = 1 To cRow
        If .Cells(i, 1).Value = ID Then get_UpdateRow = i: Exit For
    Next
End With
End Function


'########################
' Ư�� ��Ʈ�� DB ������ �迭�� ��ȯ (�̹� �������Ͽ����� ���)
' Array = Get_DB(Sheet1)
'########################
Function Get_DB(WS As Worksheet, Optional NoID As Boolean = False, Optional IncludeHeader As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim offCol As Long

If NoID = False Then offCol = -1

With WS
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS, offCol)
    Get_DB = .Range(.Cells(2 + Sgn(IncludeHeader), 1), .Cells(cRow, cCol))
End With
    
End Function
'########################
'Ư�� ��Ʈ���� ������ ID�� �ʵ� �� ��ȯ (�̹� �������� ����)
' Value = Get_Records(Sheet1, ID, "�ʵ��")
'########################
Function Get_Records(WS As Worksheet, ID, fields)

Dim cRow As Long: Dim cCol As Long
Dim vFields As Variant: Dim vField As Variant
Dim vFieldNo As Variant
Dim i As Long: Dim j As Long

cRow = Get_InsertRow(WS) - 1
cCol = Get_ColumnCnt(WS)

If InStr(1, fields, ",") > 0 Then vFields = Split(fields, ",") Else vFields = Array(fields)
ReDim vFieldNo(0 To UBound(vFields))

With WS
    For Each vField In vFields
        For i = 1 To cCol
            If .Cells(1, i).Value = Trim(vField) Then vFieldNo(j) = i: j = j + 1
        Next
    Next
Stop
    For i = 2 To cRow
        If .Cells(i, 1).Value = ID Then
            For j = 0 To UBound(vFieldNo)
                vFieldNo(j) = .Cells(i, vFieldNo(j))
            Next
            Exit For
        End If
    Next
    
Get_Records = vFieldNo

End With

End Function

'########################
' hjlee 2021.08.18 �߰�
'Ư�� ��Ʈ���� ������ ID�� ��ü �ʵ� �� ��ȯ
' Value = Get_Record_Array(Sheet1, ID)
'########################
Function Get_Record_Array(WS As Worksheet, ID)

    Dim cRow, cCol As Long
    Dim row, col As Long
    Dim fields As Variant
    
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS)
    
    ReDim fields(1 To cCol)

    With WS
        For row = 2 To cRow
            If .Cells(row, 1).Value = ID Then
                For col = 1 To cCol
                    fields(col) = .Cells(row, col)
                Next
                Exit For
            End If
        Next
    End With
    
    Get_Record_Array = fields

End Function

'########################
' ��Ʈ�� ���ο� ���ڵ� �߰� (�ݵ�� ù��° ���� ID, ������ �� ������� �Է�)
' Insert_Record Sheet1, ID, �ʵ�1, �ʵ�2, �ʵ�3, ..
'########################
Sub Insert_Record(WS As Worksheet, ParamArray vaParamArr() As Variant)

Dim cID As Long
Dim cRow As Long
Dim vaArr As Variant: Dim i As Long: i = 2

With WS
    cRow = Get_InsertRow(WS)
    If InStr(1, .Cells(1, 1).Value, "ID") > 0 Then
        cID = Get_MaxID(WS)
        .Cells(cRow, 1).Value = cID
        For Each vaArr In vaParamArr
            .Cells(cRow, i).Value = vaArr
            i = i + 1
        Next
    Else
        For Each vaArr In vaParamArr
            .Cells(cRow, i - 1).Value = vaArr
            i = i + 1
        Next
    End If
    
End With

End Sub
'########################
' ��Ʈ���� ID �� ���� ���ڵ��� ��� �� ������Ʈ (�ݵ�� ù��° ���� ID���� �ϸ�, ������ ���� ������� �Է�)
' Update_Record Sheet1, ID, �ʵ�1, �ʵ�2, �ʵ�3, ...
'########################
Sub Update_Record(WS As Worksheet, ParamArray vaParamArr() As Variant)

Dim cRow As Long
Dim i As Long
Dim ID As Variant

If IsNumeric(vaParamArr(0)) = True Then ID = CLng(vaParamArr(0)) Else ID = vaParamArr(0)

With WS
    cRow = get_UpdateRow(WS, ID)
    
    For i = 1 To UBound(vaParamArr)
        If Not IsMissing(vaParamArr(i)) Then .Cells(cRow, i + 1).Value = vaParamArr(i)
    Next
    
End With
End Sub

'########################
' hjlee. 2021.08.18 �߰�
' ��Ʈ���� �ʵ���� �÷� ��ȣ�� ����
' Get_Column_Index Sheet1, "�ʵ��"
'########################
Function Get_Column_Index(WS As Worksheet, vFieldName) As Long

Dim i As Long
Dim cCol As Long

With WS
    cCol = Get_ColumnCnt(WS)
    For i = 1 To cCol
        If .Cells(1, i).Value = vFieldName Then Get_Column_Index = i: Exit For
    Next
End With

End Function


'########################
' hjlee. 2021.08.18 �߰�
' ��Ʈ���� ID �� ���� ���ڵ��� vFieldName �ʵ尪�� vData�� ������Ʈ
' Update_Record_Column Sheet1, ID, "�÷���", "������ ��"
'########################
Sub Update_Record_Column(WS As Worksheet, ID, vFieldName, vData As Variant)

Dim cRow As Long
Dim cCol As Long

If IsNumeric(ID) = True Then ID = CLng(ID)

With WS
    cRow = get_UpdateRow(WS, ID)
    cCol = Get_Column_Index(WS, vFieldName)
    .Cells(cRow, cCol).Value = vData
End With

End Sub
'########################
' ��Ʈ���� ID �� ���� ���ڵ� ����
' Delete_Record Sheet1, ID
'########################
Sub Delete_Record(WS As Worksheet, ID)

Dim cRow As Long

If IsNumeric(ID) = True Then ID = CLng(ID)

With WS
    cRow = get_UpdateRow(WS, ID)
    .Cells(cRow, 1).EntireRow.Delete
End With

End Sub

'########################
' �迭�� �ܺ�IDŰ �ʵ带 �� ��ƮDB�� �����Ͽ� �ش� �ܺ�IDŰ�� ������ ���� �迭�� ��ȯ
' Array = Connect_DB(Get_DB(Sheet1),2,Sheet2, "�ʵ�1, �ʵ�2, �ʵ�3")
'########################
Function Connect_DB(db As Variant, ForeignID_Fields As Variant, FromWS As Worksheet, fields As String, Optional IncludeHeader As Boolean = False)

Dim cRow As Long: Dim cCol As Long
Dim vForeignID_Fields As Variant: Dim vForeignID_Field As Variant
Dim ForeignID As Variant
Dim vFields As Variant: Dim vField As Variant
Dim vID As Variant: Dim vFieldNo As Variant
Dim Dict As Object
Dim i As Long: Dim j As Long
Dim AddCols As Long


cRow = UBound(db, 1)
cCol = UBound(db, 2)
If InStr(1, fields, ",") > 1 Then
    AddCols = Len(fields) - Len(Replace(fields, ",", "")) + 1
    vFields = Split(fields, ",")
Else
    AddCols = 1
    vFields = Array(fields)
End If

ReDim Preserve db(1 To cRow, 1 To cCol + AddCols)
        
Set Dict = Get_Dict(FromWS)
vID = Dict("ID")

ReDim vFieldNo(0 To UBound(vFields))

For Each vField In vFields
    For i = 1 To UBound(vID)
        If vID(i) = Trim(vField) Then vFieldNo(j) = i: j = j + 1
    Next
Next

If InStr(1, ForeignID_Fields, ",") > 0 Then vForeignID_Fields = Split(ForeignID_Fields, ",") Else vForeignID_Fields = Array(ForeignID_Fields)

For Each vForeignID_Field In vForeignID_Fields
    For i = 1 To cRow
        If IncludeHeader = True And i = 1 Then ForeignID = "ID" Else ForeignID = db(i, Trim(vForeignID_Field))
        If Dict.Exists(ForeignID) Then
            For j = 1 To AddCols
                db(i, cCol + j) = Dict(ForeignID)(vFieldNo(j - 1))
            Next
        End If
    Next
Next

Connect_DB = db
    
End Function

'########################
' hjlee 2021.08.23 �߰�
' �迭�� �ܺ�IDŰ �ʵ带 �� ��ƮDB�� �����Ͽ� �ش� �ܺ�IDŰ�� ������ ���� �迭�� ��ȯ
' Array = Join_DB(Get_DB(Sheet1), 2, Sheet2, "JOIN�ʵ�", "�����ʵ�1, �����ʵ�2, �����ʵ�3")
'########################
Function Join_DB(db As Variant, ForeignID_Fields As Variant, FromWS As Worksheet, joinField As String, returnFields As String, Optional IncludeHeader As Boolean = False)

Dim cRow As Long: Dim cCol As Long
Dim vForeignID_Fields As Variant: Dim vForeignID_Field As Variant
Dim ForeignID As Variant
Dim vFields As Variant: Dim vField As Variant
Dim vID As Variant: Dim vFieldNo As Variant
Dim Dict As Object
Dim i As Long: Dim j As Long
Dim AddCols As Long


cRow = UBound(db, 1)
cCol = UBound(db, 2)
If InStr(1, returnFields, ",") > 1 Then
    AddCols = Len(returnFields) - Len(Replace(returnFields, ",", "")) + 1
    vFields = Split(returnFields, ",")
Else
    AddCols = 1
    vFields = Array(returnFields)
End If

ReDim Preserve db(1 To cRow, 1 To cCol + AddCols)
        
Set Dict = Get_Dict_KeyField(FromWS, joinField)
vID = Dict(joinField)

ReDim vFieldNo(0 To UBound(vFields))

For Each vField In vFields
    For i = 1 To UBound(vID)
        If vID(i) = Trim(vField) Then vFieldNo(j) = i: j = j + 1
    Next
Next

If InStr(1, ForeignID_Fields, ",") > 0 Then vForeignID_Fields = Split(ForeignID_Fields, ",") Else vForeignID_Fields = Array(ForeignID_Fields)

For Each vForeignID_Field In vForeignID_Fields
    For i = 1 To cRow
        If IncludeHeader = True And i = 1 Then ForeignID = joinField Else ForeignID = db(i, Trim(vForeignID_Field))
        If ForeignID <> "" Then
            If Dict.Exists(CLng(ForeignID)) Then
                For j = 1 To AddCols
                    db(i, cCol + j) = Dict(CLng(ForeignID))(vFieldNo(j - 1))
                Next
            End If
        End If
    Next
Next

Join_DB = db
    
End Function

'########################
' Ư�� �迭���� Value�� �����ϴ� ���ڵ常 ã�� �ٽ� �迭�� ��ȯ
' Array = Filtered_DB(Array, "�˻���", False)
'########################
Function Filtered_DB(db, Value, Optional FilterCol, Optional ExactMatch As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim vArr As Variant: Dim s As String: Dim filterArr As Variant:  Dim Cols As Variant: Dim col As Variant: Dim Colcnt As Long
Dim isDateVal As Boolean
Dim vReturn As Variant: Dim vResult As Variant
Dim Dict As Object: Dim dictKey As Variant
Dim i As Long: Dim j As Long
Dim Operator As String

Set Dict = CreateObject("Scripting.Dictionary")

If Value <> "" Then
    cRow = UBound(db, 1)
    cCol = UBound(db, 2)
    ReDim vArr(1 To cRow)
    For i = 1 To cRow
        s = ""
        For j = 1 To cCol
            s = s & db(i, j) & "|^"
        Next
        vArr(i) = s
    Next
    
    If IsMissing(FilterCol) Then
        filterArr = vArr
    Else
        Cols = Split(FilterCol, ",")
        ReDim filterArr(1 To cRow)
        For i = 1 To cRow
            s = ""
            For Each col In Cols
                s = s & db(i, Trim(col)) & "|^"
            Next
            filterArr(i) = s
        Next
    End If
    
    If Left(Value, 2) = ">=" Or Left(Value, 2) = "<=" Or Left(Value, 2) = "=>" Or Left(Value, 2) = "=<" Then
        Operator = Left(Value, 2)
        If IsDate(Right(Value, Len(Value) - 2)) Then isDateVal = True
    ElseIf Left(Value, 1) = ">" Or Left(Value, 1) = "<" Then
        Operator = Left(Value, 1)
        If IsDate(Right(Value, Len(Value) - 1)) Then isDateVal = True
    Else: End If
    
    If Operator <> "" Then
        If isDateVal = False Then
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) > CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) < CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) >= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) <= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
            End Select
        Else
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) > CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) < CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) >= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) <= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
            End Select
        End If
    Else
        If ExactMatch = False Then
            For i = 1 To cRow
                If filterArr(i) Like "*" & Value & "*" Then
                    vArr(i) = Left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    Dict.Add i, vReturn
                End If
            Next
        Else
            For i = 1 To cRow
                If filterArr(i) Like Value & "|^" Then
                    vArr(i) = Left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    Dict.Add i, vReturn
                End If
            Next
        End If
    End If
        
    If Dict.Count > 0 Then
        ReDim vResult(1 To Dict.Count, 1 To cCol)
        i = 1
        For Each dictKey In Dict.Keys
            For j = 1 To cCol
                vResult(i, j) = Dict(dictKey)(j - 1)
            Next
            i = i + 1
        Next
    End If
    
    Filtered_DB = vResult
Else
    Filtered_DB = db
End If

End Function

'########################
' �� ��ǰ�� �ܰ������ ����մϴ�.
' DB = Get_Balance(DB, shtInventory, �԰��������ȣ, ����������ȣ, ��ǰID����ȣ)
'########################

Function Get_Balance(db, InventoryWS As Worksheet, ColumnIN, ColumnOUT, ColumnID) As Variant

Dim InventoryDB As Variant
Dim Dict As Dictionary
Dim cRow As Long: Dim cCol As Long
Dim i As Long: Dim cID

If Not IsNumeric(ColumnOUT) Then ColumnOUT = Range(ColumnOUT & 1).Column
If Not IsNumeric(ColumnIN) Then ColumnIN = Range(ColumnIN & 1).Column
If Not IsNumeric(ColumnID) Then ColumnID = Range(ColumnID & 1).Column

cRow = UBound(db, 1)
cCol = UBound(db, 2)
Set Dict = CreateObject("Scripting.Dictionary")

ReDim Preserve db(1 To cRow, 1 To cCol + 1)

For i = 1 To cRow:    Dict.Add db(i, 1), 0: Next
InventoryDB = Get_DB(InventoryWS)

For i = LBound(InventoryDB, 1) To UBound(InventoryDB, 1)
    cID = InventoryDB(i, ColumnID)
    If Dict.Exists(cID) Then
        Dict(cID) = Dict(cID) + InventoryDB(i, CLng(ColumnIN)) - InventoryDB(i, CLng(ColumnOUT))
    End If
Next

For i = LBound(db, 1) To UBound(db, 1)
    db(i, cCol + 1) = Dict(db(i, 1))
Next

Get_Balance = db

End Function

'########################
' Ư�� ��Ʈ�� DB ������ Dictionary�� ��ȯ (�̹� �������Ͽ����� ���)
' Dict = GetDict(Sheet1)
'########################
Function Get_Dict(WS As Worksheet) As Object

Dim cRow As Long: Dim cCol As Long
Dim Dict As Object
Dim vArr As Variant
Dim i As Long: Dim j As Long

Set Dict = CreateObject("Scripting.Dictionary")

With WS
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS)
    
    For i = 1 To cRow
        ReDim vArr(1 To cCol - 1)
        For j = 2 To cCol
            vArr(j - 1) = .Cells(i, j)
        Next
        Dict.Add .Cells(i, 1).Value, vArr
    Next
End With

Set Get_Dict = Dict

End Function

'########################
' hjlee 2021.08.24 �߰�
' Ư�� ��Ʈ�� DB ������ Dictionary�� ��ȯ (�̹� �������Ͽ����� ���)
' keyFieldName�� �������� Dict ����
' Dict = Get_Dict_KeyField(Sheet1, keyFieldName as string)
'########################
Function Get_Dict_KeyField(WS As Worksheet, keyFieldName As String) As Object

Dim cRow As Long: Dim cCol As Long
Dim Dict As Object
Dim vArr As Variant
Dim i As Long: Dim j As Long
Dim keyFieldNo As Long

Set Dict = CreateObject("Scripting.Dictionary")

With WS
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS)
    
    keyFieldNo = 1
    For i = 1 To cCol
        If .Cells(1, i) = keyFieldName Then
            keyFieldNo = i
        End If
    Next
    
    For i = 1 To cRow
        ReDim vArr(1 To cCol)
        For j = 1 To cCol
            vArr(j) = .Cells(i, j)
        Next
        
        'Dict(.Cells(i, keyFieldNo).Value) = vArr
        If Dict.Exists(.Cells(i, keyFieldNo).Value) Then
            Dict.Remove (.Cells(i, keyFieldNo).Value)
        End If
        Dict.Add .Cells(i, keyFieldNo).Value, vArr
    Next
End With

Set Get_Dict_KeyField = Dict

End Function

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Arr_To_Dict �Լ�
'�� ������ Dictionary �� ��ȯ�մϴ�.
'�� �μ� ����
'_____________Arr       : Dictionary�� ��ȯ�� �迭�Դϴ�.
'�� ��� ����
'Dict = Arr_To_Dict(Arr)
'##############################################################
Function Arr_To_Dict(arr As Variant) As Object

Dim Dict As Object: Dim vArr As Variant
Dim cCol As Long
Dim i As Long: Dim j As Long

Set Dict = CreateObject("Scripting.Dictionary")
cCol = UBound(arr, 2)

For i = LBound(arr, 1) To UBound(arr, 1)
        ReDim vArr(1 To cCol - 1)
        For j = 2 To cCol
            vArr(j - 1) = arr(i, j)
        Next
        Dict.Add arr(i, 1), vArr
Next

Set Arr_To_Dict = Dict

End Function

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Dict_To_Arr �Լ�
'�� Dictionary�� ������ ��ȯ�մϴ�.
'�� �μ� ����
'_____________Dict       : �迭�� ��ȯ�� Dictionary �Դϴ�.
'�� ��� ����
'Arr = Dict_To_Arr(Dict)
'##############################################################
Function Dict_To_Arr(Dict As Object) As Variant

Dim i As Long: Dim j As Long: Dim dictKey As Variant: Dim cCol As Long
Dim vTest As Variant
i = 1

If Dict.Count > 0 Then
    If IsObject(Dict(Dict.Keys()(0))) Then cCol = UBound(Dict(Dict.Keys()(0))) Else cCol = 1
    ReDim vResult(1 To Dict.Count, 1 To cCol + 1)
    For Each dictKey In Dict.Keys
        vResult(i, 1) = dictKey
        If cCol = 1 Then
            vResult(i, 2) = Dict(dictKey)
        Else
            For j = 2 To cCol + 1
                vResult(i, j) = Dict(dictKey)(j - 1)
            Next
        End If
        i = i + 1
    Next
End If

Dict_To_Arr = vResult
    
End Function
'########################
' ��Ʈ�� Ư�� �ʵ� ������ �߰��Ǵ� ���� ���������� Ȯ��. �������� ��� TRUE�� ��ȯ
' boolean = IsUnique(Sheet1, "���", 1)
'########################
Function IsUnique(db As Variant, uniqueVal, Optional ColNo As Long = 1, Optional Exclude) As Boolean

Dim endRow As Long
Dim i As Long

For i = LBound(db, 1) To UBound(db, 1)
    If db(i, ColNo) = uniqueVal Then
        If Not IsMissing(Exclude) Then
            If Exclude <> uniqueVal Then
                IsUnique = False
                Exit Function
            End If
        Else
            IsUnique = False: Exit Function
        End If
    End If
Next

IsUnique = True

End Function

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Extract_Column �Լ�
'�� �迭���� ������ ���� �����մϴ�.
'�� �μ� ����
'_____________DB        : Ư�� ���� ������ �迭�Դϴ�.
'_____________Col       : �迭���� ������ ���� ����ȣ�Դϴ�.
'�� ��� ����
'Arr = Extract_Column(Arr, 3) '<- 3��° ���� �����մϴ�.
'##############################################################

Function Extract_Column(db As Variant, col As Long) As Variant

Dim i As Long
Dim vArr As Variant

ReDim vArr(LBound(db) To UBound(db), 1 To 1)
For i = LBound(db) To UBound(db)
        vArr(i, 1) = db(i, col)
Next

Extract_Column = vArr

End Function
