Attribute VB_Name = "Mod_ImportManage"
Option Explicit

'관리 구조체
Type Manage
    ID As Long
    수입지출 As String
    분류1 As String
    분류2 As String
    관리번호 As String
    거래처 As String
    품목 As String
    재질 As String
    규격 As String
    단가 As String
    금액 As String
    단위 As String
    중량 As String
    수량 As String
    수주 As String
    납기 As String
    발주 As String
    입고 As String
    납품 As String
    명세서 As String
    계산서 As String
    결재 As String
    결재월 As String
    부가세 As String
    등록일자 As String
    수정일자 As String
End Type

Sub ImportManage()

    Dim WB As Workbook
    Dim WS As Worksheet:
    Dim i As Long
    Dim j As Long
    Dim strWS As String
    Dim manageFileList(1) As Variant
    Dim importCount As Long
    Dim pos As Long
    Dim M As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ClearManageData
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2005관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 5 And M <= 6 Then
                        importCount = ImportManageData_Type9(WS, WS.Name, 2005)
                    Else
                        importCount = ImportManageData_Type8(WS, WS.Name, 2005)
                    End If
                End If
            End If
        End If
    Next
    WB.Close

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2006관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 2 Then
                        importCount = ImportManageData_Type7(WS, WS.Name, 2006)
                    Else
                        importCount = ImportManageData_Type6(WS, WS.Name, 2006)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2007관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 3 Then
                        importCount = ImportManageData_Type6(WS, WS.Name, 2007)
                    Else
                        importCount = ImportManageData_Type5(WS, WS.Name, 2007)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2008관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2008)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2009관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2009)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2010관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type5(WS, WS.Name, 2010)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2011관리.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If IsNumeric(M) Then
                    If M >= 1 And M <= 2 Then
                        importCount = ImportManageData_Type4(WS, WS.Name, 2011)
                    Else
                        importCount = ImportManageData_Type3(WS, WS.Name, 2011)
                    End If
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2012관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type3(WS, WS.Name, 2012)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\2013관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type2(WS, WS.Name, 2013)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2013관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If M >= 1 And M <= 10 Then
                    importCount = ImportManageData_Type2(WS, WS.Name, 2013)
                Else
                    importCount = ImportManageData_Type1(WS, WS.Name, 2013)
                End If
            End If
            
        End If
    Next
    WB.Close


    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2014관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2014)
        End If
    Next
    WB.Close

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2015관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2015)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2016관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2016)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2017관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2017)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2018관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2018)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2019관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2019)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2020관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2020)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-관리문서\법인-2021관리.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportManageData_Type1(WS, WS.Name, 2021)
        End If
    Next
    WB.Close

End Sub

Sub ClearManageData()
    Dim endCol, endRow As Long
    
    With shtManageData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtManageMemoData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub


Function ImportManageData_Type1(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    With WS
        endCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            man.분류1 = .Cells(i, 2)
            man.분류2 = .Cells(i, 3)
            man.관리번호 = .Cells(i, 4)
            man.거래처 = .Cells(i, 5)
            man.품목 = .Cells(i, 6)
            man.재질 = .Cells(i, 7)
            man.규격 = .Cells(i, 8)
            man.단가 = .Cells(i, 9)
            man.금액 = .Cells(i, 10)
            man.단위 = .Cells(i, 11)
            man.중량 = .Cells(i, 12)
            man.수량 = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.수주 = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.납기 = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.납기 = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.발주 = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.입고 = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.납품 = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.계산서 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.결재 = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 30)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            man.부가세 = .Cells(i, 31)
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type1 = importCount
End Function

Function ImportManageData_Type2(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            man.분류1 = .Cells(i, 2)
            man.분류2 = .Cells(i, 3)
            man.관리번호 = .Cells(i, 4)
            man.거래처 = .Cells(i, 5)
            man.품목 = .Cells(i, 6)
            man.재질 = .Cells(i, 7)
            man.규격 = .Cells(i, 8)
            man.단가 = .Cells(i, 9)
            man.금액 = .Cells(i, 10)
            man.단위 = .Cells(i, 11)
            man.중량 = .Cells(i, 12)
            man.수량 = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.수주 = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.납기 = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.납기 = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.발주 = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.입고 = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.납품 = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.계산서 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.결재 = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 30)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            man.부가세 = .Cells(i, 31)
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type2 = importCount
End Function

Function ImportManageData_Type3(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        Else
            regDate = DateSerial(Y, M, 1)
        End If
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            man.분류1 = .Cells(i, 2)
            man.분류2 = .Cells(i, 3)
            man.관리번호 = .Cells(i, 4)
            man.거래처 = .Cells(i, 5)
            man.품목 = .Cells(i, 6)
            man.재질 = .Cells(i, 7)
            man.규격 = .Cells(i, 8)
            man.단가 = .Cells(i, 9)
            man.금액 = .Cells(i, 10)
            man.단위 = .Cells(i, 11)
            man.중량 = .Cells(i, 12)
            man.수량 = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.수주 = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.납기 = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.납기 = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.발주 = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.입고 = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.납품 = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.계산서 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.결재 = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 30)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
                
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type3 = importCount
End Function

Function ImportManageData_Type4(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음 - 데이터 시트에 열을 하나 추가해서 해결함
            man.분류1 = .Cells(i, 2)
            man.분류2 = .Cells(i, 3)
            man.관리번호 = .Cells(i, 4)
            man.거래처 = .Cells(i, 5)
            man.품목 = .Cells(i, 6)
            man.재질 = .Cells(i, 7)
            man.규격 = .Cells(i, 8)
            man.단가 = .Cells(i, 9)
            man.금액 = .Cells(i, 10)
            man.단위 = .Cells(i, 11)
            man.중량 = .Cells(i, 12)
            man.수량 = .Cells(i, 13)
            If .Cells(i, 14) <> "" And .Cells(i, 15) <> "" And IsNumeric(.Cells(i, 14)) And IsNumeric(.Cells(i, 15)) Then
                man.수주 = DateSerial(Y, .Cells(i, 14), .Cells(i, 15))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 16) <> "" And .Cells(i, 17) <> "" And IsNumeric(.Cells(i, 16)) And IsNumeric(.Cells(i, 17)) Then
                man.납기 = DateSerial(Y, .Cells(i, 16), .Cells(i, 17))
            Else
                man.납기 = ""
            End If
            If .Cells(i, 18) <> "" And .Cells(i, 19) <> "" And IsNumeric(.Cells(i, 18)) And IsNumeric(.Cells(i, 19)) Then
                man.발주 = DateSerial(Y, .Cells(i, 18), .Cells(i, 19))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 20) <> "" And .Cells(i, 21) <> "" And IsNumeric(.Cells(i, 20)) And IsNumeric(.Cells(i, 21)) Then
                man.입고 = DateSerial(Y, .Cells(i, 20), .Cells(i, 21))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 22) <> "" And .Cells(i, 23) <> "" And IsNumeric(.Cells(i, 22)) And IsNumeric(.Cells(i, 23)) Then
                man.납품 = DateSerial(Y, .Cells(i, 22), .Cells(i, 23))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 24) <> "" And .Cells(i, 25) <> "" And IsNumeric(.Cells(i, 24)) And IsNumeric(.Cells(i, 25)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 24), .Cells(i, 25))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 26) <> "" And .Cells(i, 27) <> "" And IsNumeric(.Cells(i, 26)) And IsNumeric(.Cells(i, 27)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 26), .Cells(i, 27))
            Else
                man.계산서 = ""
            End If
                
            If .Cells(i, 28) <> "" And .Cells(i, 29) <> "" And IsNumeric(.Cells(i, 28)) And IsNumeric(.Cells(i, 29)) Then
                man.결재 = DateSerial(Y, .Cells(i, 28), .Cells(i, 29))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 30)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type4 = importCount
End Function

Function ImportManageData_Type5(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음
            man.분류1 = ""
            man.분류2 = .Cells(i, 2)
            man.관리번호 = .Cells(i, 3)
            man.거래처 = .Cells(i, 4)
            man.품목 = .Cells(i, 5)
            man.재질 = .Cells(i, 6)
            man.규격 = .Cells(i, 7)
            man.단가 = .Cells(i, 8)
            man.금액 = .Cells(i, 9)
            man.단위 = .Cells(i, 10)
            man.중량 = .Cells(i, 11)
            man.수량 = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.수주 = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.납기 = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.납기 = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.발주 = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.입고 = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.납품 = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.계산서 = ""
            End If
                
            If .Cells(i, 27) <> "" And .Cells(i, 28) <> "" And IsNumeric(.Cells(i, 27)) And IsNumeric(.Cells(i, 28)) Then
                man.결재 = DateSerial(Y, .Cells(i, 27), .Cells(i, 28))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 29)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type5 = importCount
End Function

Function ImportManageData_Type6(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음
            man.분류1 = ""
            man.분류2 = .Cells(i, 2)
            man.관리번호 = .Cells(i, 3)
            man.거래처 = .Cells(i, 4)
            man.품목 = .Cells(i, 5)
            man.재질 = .Cells(i, 6)
            man.규격 = .Cells(i, 7)
            man.단가 = .Cells(i, 8)
            man.금액 = .Cells(i, 9)
            man.단위 = .Cells(i, 10)
            man.중량 = .Cells(i, 11)
            man.수량 = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.수주 = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.발주 = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.입고 = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.납품 = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.계산서 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.결재 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.결재 = ""
            End If
            
            man.결재월 = .Cells(i, 27)
            If man.결재월 <> "" Then
                If IsNumeric(man.결재월) Then
                    M = CLng(man.결재월)
                    If M >= 1 And M <= 12 Then
                        man.결재월 = DateSerial(Y, M, 1)
                    End If
                End If
            End If
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type6 = importCount
End Function

Function ImportManageData_Type7(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type2는 머리글 9행, 데이터 11행으로 분석
    With WS
        endCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 11 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음
            man.분류1 = ""
            man.분류2 = .Cells(i, 2)
            man.관리번호 = .Cells(i, 3)
            man.거래처 = .Cells(i, 4)
            man.품목 = .Cells(i, 5)
            man.재질 = .Cells(i, 6)
            man.규격 = .Cells(i, 7)
            man.단가 = .Cells(i, 8)
            man.금액 = .Cells(i, 9)
            man.단위 = .Cells(i, 10)
            man.중량 = .Cells(i, 11)
            man.수량 = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.수주 = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.발주 = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.입고 = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.납품 = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.계산서 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.결재 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.결재 = ""
            End If
            
            'Type7은 결재월 없음
            man.결재월 = ""
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type7 = importCount
End Function

Function ImportManageData_Type8(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type8은 머리글 8행, 데이터 10행으로 분석
    With WS
        endCol = .Cells(8, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 10 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음
            man.분류1 = ""
            man.분류2 = .Cells(i, 2)
            man.관리번호 = .Cells(i, 3)
            man.거래처 = .Cells(i, 4)
            man.품목 = .Cells(i, 5)
            man.재질 = .Cells(i, 6)
            man.규격 = .Cells(i, 7)
            man.단가 = .Cells(i, 8)
            man.금액 = .Cells(i, 9)
            man.단위 = .Cells(i, 10)
            man.중량 = .Cells(i, 11)
            man.수량 = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.수주 = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.발주 = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.입고 = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.납품 = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.계산서 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.결재 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.결재 = ""
            End If
            
            'Type7은 결재월 없음
            man.결재월 = ""
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
        
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type8 = importCount
End Function

Function ImportManageData_Type9(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim man As Manage
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M, D As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        If Not IsNumeric(M) Then
            Exit Function
        End If
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        Exit Function
    End If
    
    'Type9는 머리글 3행, 데이터 5행으로 분석
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 5 To endRow
 
        With WS
            man.수입지출 = .Cells(i, 1)
            'Type4는 분류1 없음
            man.분류1 = ""
            man.분류2 = .Cells(i, 2)
            man.관리번호 = .Cells(i, 3)
            man.거래처 = .Cells(i, 4)
            man.품목 = .Cells(i, 5)
            man.재질 = .Cells(i, 6)
            man.규격 = .Cells(i, 7)
            man.단가 = .Cells(i, 8)
            man.금액 = .Cells(i, 9)
            man.단위 = .Cells(i, 10)
            man.중량 = .Cells(i, 11)
            man.수량 = .Cells(i, 12)
            If .Cells(i, 13) <> "" And .Cells(i, 14) <> "" And IsNumeric(.Cells(i, 13)) And IsNumeric(.Cells(i, 14)) Then
                man.수주 = DateSerial(Y, .Cells(i, 13), .Cells(i, 14))
            Else
                man.수주 = ""
            End If
            If .Cells(i, 15) <> "" And .Cells(i, 16) <> "" And IsNumeric(.Cells(i, 15)) And IsNumeric(.Cells(i, 16)) Then
                man.발주 = DateSerial(Y, .Cells(i, 15), .Cells(i, 16))
            Else
                man.발주 = ""
            End If
            If .Cells(i, 17) <> "" And .Cells(i, 18) <> "" And IsNumeric(.Cells(i, 17)) And IsNumeric(.Cells(i, 18)) Then
                man.입고 = DateSerial(Y, .Cells(i, 17), .Cells(i, 18))
            Else
                man.입고 = ""
            End If
            If .Cells(i, 19) <> "" And .Cells(i, 20) <> "" And IsNumeric(.Cells(i, 19)) And IsNumeric(.Cells(i, 20)) Then
                man.납품 = DateSerial(Y, .Cells(i, 19), .Cells(i, 20))
            Else
                man.납품 = ""
            End If
            If .Cells(i, 21) <> "" And .Cells(i, 22) <> "" And IsNumeric(.Cells(i, 21)) And IsNumeric(.Cells(i, 22)) Then
                man.명세서 = DateSerial(Y, .Cells(i, 21), .Cells(i, 22))
            Else
                man.명세서 = ""
            End If
            If .Cells(i, 23) <> "" And .Cells(i, 24) <> "" And IsNumeric(.Cells(i, 23)) And IsNumeric(.Cells(i, 24)) Then
                man.계산서 = DateSerial(Y, .Cells(i, 23), .Cells(i, 24))
            Else
                man.계산서 = ""
            End If
            If .Cells(i, 25) <> "" And .Cells(i, 26) <> "" And IsNumeric(.Cells(i, 25)) And IsNumeric(.Cells(i, 26)) Then
                man.결재 = DateSerial(Y, .Cells(i, 25), .Cells(i, 26))
            Else
                man.결재 = ""
            End If
            
            'Type7은 결재월 없음
            man.결재월 = ""
            
            'Type3는 부가세 없음
            man.부가세 = ""
            man.등록일자 = regDate
        End With
        
        '관리 테이블 등록
        currentId = Get_CurrentID(shtManageData)
        Insert_Record shtManageData, man.수입지출, man.분류1, man.분류2, man.관리번호, man.거래처, man.품목, _
                      man.재질, man.규격, man.단가, man.금액, man.단위, man.중량, man.수량, man.수주, man.납기, man.발주, man.입고, _
                      man.납품, man.명세서, man.계산서, man.결재, man.결재월, man.부가세, man.등록일자, man.수정일자
        
        importCount = importCount + 1
        
       
        '메모 테이블 등록
        For j = 1 To endCol
            If Not WS.Cells(i, j).Comment Is Nothing Then
                Insert_Record shtManageMemoData, currentId, man.관리번호, WS.Cells(i, j).Comment.Text, regDate
            End If
        Next
    Next
    
    ImportManageData_Type9 = importCount
End Function


