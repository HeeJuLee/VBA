Attribute VB_Name = "z_Mod_Sort"
Option Explicit

Function Sort2DArray(db, ByVal Index As Long, Optional ByVal order As Integer = -1, Optional ByVal ByColumn As Boolean = False, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = 0, Optional THRESHOLD As Long = 20)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'���� �� ���� �� ��ó�� �ݵ�� ����ؾ� �մϴ�.
'
'�� Sort2DArray ��ɹ�
'�� 2���� �迭�� ��������/�������� �Ǵ� ������������ �����մϴ�. �Ѱ����� �����Ͽ� QuickSort �Ǵ� InsertionSort�� ���� ������ ������ �� �ֽ��ϴ�. �⺻ �Ѱ����� 20�Դϴ�.
'�� �μ� ����
'_____________DB                : ������ �迭�Դϴ�.
'_____________Index            : DB�� ������ ���� �����Դϴ�.
'_____________Order           : [�����μ�] 1 �̸� �������� �����մϴ�. �⺻���� -1 (=��������) �����Դϴ�.
'_____________ByColumn    : [�����μ�] True �� ������(���ι���) �����Դϴ�. �⺻���� FALSE �Դϴ�.
'_____________lngStart        : [�����μ�] ������ ������ �������Դϴ�. �⺻���� �迭�� ���������Դϴ�.
'_____________lngEnd         : [�����μ�] ������ ������ ���������Դϴ�. �⺻���� �迭�� �����������Դϴ�.
'_____________Threshold    : [�����μ�] QuickSort�� InsertionSort�� ������ �Ѱ����Դϴ�. �⺻���� 20 �Դϴ�. Threshold �� ���Ǵ� �������� ������ ���� �ٸ����� ��κ� 10-20 ������ ������ ���˴ϴ�.
'�� ��ȯ��
'_____________���ĵ� �迭�� ��ȯ�մϴ�.
'�� ��ɹ��� �Ʒ� ��ũ�� �����Ͽ� �ۼ��� ��ɹ��Դϴ�.
'https://www.vbforums.com/showthread.php?631366-RESOLVED-Quick-Sort-2D-Array
'###############################################################

Dim i As Long: Dim j As Long: Dim k As Long
Dim Pivot: Dim Temp
Dim Stack(1 To 64) As Long: Dim StackPtr As Long

If lngStart = 0 Then
    If ByColumn = False Then lngStart = LBound(db, 1) Else lngStart = LBound(db, 2)
End If

If lngEnd = 0 Then
    If ByColumn = False Then lngEnd = UBound(db, 1) Else lngEnd = UBound(db, 2)
End If

'���ι��� ����
  If ByColumn Then
    ReDim Temp(LBound(db, 1) To UBound(db, 1))
    Stack(StackPtr + 1) = lngStart
    Stack(StackPtr + 2) = lngEnd
    StackPtr = StackPtr + 2
    Do
      StackPtr = StackPtr - 2
      lngStart = Stack(StackPtr + 1)
      lngEnd = Stack(StackPtr + 2)
      If lngEnd - lngStart < THRESHOLD Then
        ' �� ����� ù��° ���� �������� ���̰� 20 �̸��� ��� Insertion Sort
        For j = lngStart + 1 To lngEnd
          For k = LBound(db, 1) To UBound(db, 1)
            Temp(k) = db(k, j)
          Next
          Pivot = db(Index, j)
          For i = j - 1 To lngStart Step -1
            If order >= 0 Then
              If db(Index, i) <= Pivot Then Exit For
            Else
              If db(Index, i) >= Pivot Then Exit For
            End If
            For k = LBound(db) To UBound(db)
              db(k, i + 1) = db(k, i)
            Next
          Next
          For k = LBound(db) To UBound(db)
            db(k, i + 1) = Temp(k)
          Next
        Next
      Else
        ' �� ����� ù��° ���� �������� ���̰� 20 �̻��� ��� Quick Sort
        i = lngStart: j = lngEnd
        Pivot = db(Index, (lngStart + lngEnd) \ 2)
        Do
          If order >= 0 Then
            Do While (db(Index, i) < Pivot): i = i + 1: Loop
            Do While (db(Index, j) > Pivot): j = j - 1: Loop
          Else
            Do While (db(Index, i) > Pivot): i = i + 1: Loop
            Do While (db(Index, j) < Pivot): j = j - 1: Loop
          End If
          If i <= j Then
            If i < j Then
              For k = LBound(db) To UBound(db)
                Temp(k) = db(k, i)
                db(k, i) = db(k, j)
                db(k, j) = Temp(k)
              Next
            End If
            i = i + 1: j = j - 1
          End If
        Loop Until i > j
        If (lngStart < j) Then
          Stack(StackPtr + 1) = lngStart
          Stack(StackPtr + 2) = j
          StackPtr = StackPtr + 2
        End If
        If (i < lngEnd) Then
          Stack(StackPtr + 1) = i
          Stack(StackPtr + 2) = lngEnd
          StackPtr = StackPtr + 2
        End If
      End If
    Loop Until StackPtr = 0
'���ι��� ����
Else
    ReDim Temp(LBound(db, 2) To UBound(db, 2))
        ' Stack ����
        Stack(StackPtr + 1) = lngStart
        Stack(StackPtr + 2) = lngEnd
        StackPtr = StackPtr + 2
            Do
                StackPtr = StackPtr - 2
                lngStart = Stack(StackPtr + 1)
                lngEnd = Stack(StackPtr + 2)
                    ' �� ����� ù��° ���� �������� ���̰� 20 �̸��� ��� Insertion Sort
                    If lngEnd - lngStart < THRESHOLD Then
                          For j = lngStart + 1 To lngEnd
                            For k = LBound(db, 2) To UBound(db, 2)
                              Temp(k) = db(j, k)
                            Next
                            Pivot = db(j, Index)
                            For i = j - 1 To lngStart Step -1
                              If order >= 0 Then
                                If db(i, Index) <= Pivot Then Exit For
                              Else
                                If db(i, Index) >= Pivot Then Exit For
                              End If
                              For k = LBound(db, 2) To UBound(db, 2)
                                db(i + 1, k) = db(i, k)
                              Next
                            Next
                            For k = LBound(db, 2) To UBound(db, 2)
                              db(i + 1, k) = Temp(k)
                            Next
                          Next
                Else
                    ' �� ����� ù��° ���� �������� ���̰� 20 �̻��� ��� Quick Sort
                    i = lngStart: j = lngEnd
                    Pivot = db((lngStart + lngEnd) \ 2, Index)
                        Do
                            If order >= 0 Then
                                Do While (db(i, Index) < Pivot): i = i + 1: Loop
                                Do While (db(j, Index) > Pivot): j = j - 1: Loop
                            Else
                                Do While (db(i, Index) > Pivot): i = i + 1: Loop
                                Do While (db(j, Index) < Pivot): j = j - 1: Loop
                            End If
                            If i <= j Then
                                  If i < j Then
                                        For k = LBound(db, 2) To UBound(db, 2)
                                            Temp(k) = db(i, k)
                                            db(i, k) = db(j, k)
                                            db(j, k) = Temp(k)
                                        Next
                                  End If
                                    i = i + 1: j = j - 1
                            End If
                        Loop Until i > j
                            If (lngStart < j) Then
                                  Stack(StackPtr + 1) = lngStart
                                  Stack(StackPtr + 2) = j
                                  StackPtr = StackPtr + 2
                            End If
                            If (i < lngEnd) Then
                                  Stack(StackPtr + 1) = i
                                  Stack(StackPtr + 2) = lngEnd
                                  StackPtr = StackPtr + 2
                            End If
                End If
            Loop Until StackPtr = 0
End If
  
Sort2DArray = db
  
End Function

