Attribute VB_Name = "Mod_ImportEstimate"
Option Explicit

'예상 실행항목
Type Production
    ID As String
    ID_견적 As String
    항목 As String
    비용 As String
    메모 As String
End Type

'견적 구조체
Type estimate
    ID As Long
    관리번호 As String
    자재번호 As String
    거래처 As String
    담당자 As String
    품명 As String
    규격 As String
    수량 As String
    단위 As String
    견적단가 As String
    견적금액 As String
    견적일 As String
    입찰일 As String
    수주일 As String
    납품일 As String
    증권일 As String
    자재비 As String
    미르 As String
    외주 As String
    인건비 As String
    기타 As String
    실행가 As String
    입찰금액 As String
    차액 As String
    마진율 As String
    수주금액 As String
    수주차액 As String
    등록일자 As String
    수정일자 As String
End Type

Sub ImportEstimate()

    Dim WB As Workbook
    Dim WS As Worksheet:
    Dim i As Long
    Dim j As Long
    Dim strWS As String
    Dim estimateFileList(1) As Variant
    Dim importCount As Long
    Dim pos As Long
    Dim M As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ClearEstimateData
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2005.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type8(WS, WS.Name, 2005)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2006.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type7(WS, WS.Name, 2006)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2006-2.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type6(WS, WS.Name, 2006)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2007.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If M >= 1 And M <= 4 Then
                    importCount = ImportEstimateData_Type6(WS, WS.Name, 2007)
                Else
                    importCount = ImportEstimateData_Type5(WS, WS.Name, 2007)
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2008.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If M >= 1 And M <= 6 Then
                    importCount = ImportEstimateData_Type5(WS, WS.Name, 2008)
                Else
                    importCount = ImportEstimateData_Type4(WS, WS.Name, 2008)
                End If
            End If
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2009.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type4(WS, WS.Name, 2009)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2010.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type4(WS, WS.Name, 2010)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2011.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type4(WS, WS.Name, 2011)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2012.xls")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type4(WS, WS.Name, 2012)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\견적관리2013.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type4(WS, WS.Name, 2013)
        End If
    Next
    WB.Close

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2013.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            pos = InStr(WS.Name, "월")
            If pos <> 0 Then
                M = Left(WS.Name, pos - 1)
                If M = 1 Or M = 2 Then
                    importCount = ImportEstimateData_Type3(WS, WS.Name, 2013)
                Else
                    importCount = ImportEstimateData_Type2(WS, WS.Name, 2013)
                End If
            End If
        End If
    Next
    WB.Close
    
Exit Sub

    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2014.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2014)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2015.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2015)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2016.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2016)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2017.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2017)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2018.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2018)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2019.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2019)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2020.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2020)
        End If
    Next
    WB.Close
    
    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2021.xlsx")
    For Each WS In WB.Worksheets
        If WS.Name Like "**월" Then
            importCount = ImportEstimateData_Type1(WS, WS.Name, 2021)
        End If
    Next
    WB.Close

End Sub

Sub ClearEstimateData()

    Dim endCol, endRow As Long
    
    With shtEstimateData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtProductionData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtEstimateMemoData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 1
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub

Function ImportEstimateData_Type1(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type1 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            est.입찰일 = Get_Date_Convert(.Cells(i, 12), year)
            est.수주일 = Get_Date_Convert(.Cells(i, 13), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 14), year)
            est.증권일 = Get_Date_Convert(.Cells(i, 15), year)
            
'            est.견적일 = .Cells(i, 11)
'            If Len(est.견적일) = 5 Then
'                est.견적일 = DateSerial(Y, Left(est.견적일, 2), Right(est.견적일, 2))
'            End If
'            est.입찰일 = .Cells(i, 12)
'            If Len(est.입찰일) = 5 Then
'                est.입찰일 = DateSerial(Y, Left(est.입찰일, 2), Right(est.입찰일, 2))
'            End If
'            est.수주일 = .Cells(i, 13)
'            If Len(est.수주일) = 5 Then
'                est.수주일 = DateSerial(Y, Left(est.수주일, 2), Right(est.수주일, 2))
'            End If
'            est.납품일 = .Cells(i, 14)
'            If Len(est.납품일) = 5 Then
'                est.납품일 = DateSerial(Y, Left(est.납품일, 2), Right(est.납품일, 2))
'            End If
'            est.증권일 = .Cells(i, 15)
'            If Len(est.증권일) = 5 Then
'                est.증권일 = DateSerial(Y, Left(est.증권일, 2), Right(est.증권일, 2))
'            End If
            
            '예상실행항목 등록
            If .Cells(i, 16) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "미르"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "인건비"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.항목 = "기타"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 21)
            est.입찰금액 = .Cells(i, 22)
            est.차액 = .Cells(i, 23)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 24)
            est.수주금액 = .Cells(i, 25)
            est.수주차액 = .Cells(i, 26)
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 16 Or j > 20 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type1 = importCount
End Function

Function ImportEstimateData_Type2(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type2 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            est.입찰일 = Get_Date_Convert(.Cells(i, 12), year)
            est.수주일 = Get_Date_Convert(.Cells(i, 13), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 14), year)
            est.증권일 = Get_Date_Convert(.Cells(i, 15), year)
            
            '예상실행항목 등록
            'Type2는 Type1과 항목명이 다름
            If .Cells(i, 16) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "외주(미르/대명/부성)"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주(현대기공/유진)"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "외주(근영/기타)"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.항목 = "명일(운송)"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 21)
            est.입찰금액 = .Cells(i, 22)
            est.차액 = .Cells(i, 23)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 24)
            est.수주금액 = .Cells(i, 25)
            
            'Type2는 수주차액이 없음
            est.수주차액 = ""
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 16 Or j > 20 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type2 = importCount
End Function

Function ImportEstimateData_Type3(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type3 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            est.입찰일 = Get_Date_Convert(.Cells(i, 12), year)
            est.수주일 = Get_Date_Convert(.Cells(i, 13), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 14), year)
            est.증권일 = Get_Date_Convert(.Cells(i, 15), year)
            
            '예상실행항목 등록
            'Type3는 Type1과 항목명이 다름
            If .Cells(i, 16) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "외주(미르/대명/부성)"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주(현대기공/유진)"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "외주(근영/기타)"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.항목 = "명일(운송)"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 21)
            est.입찰금액 = .Cells(i, 22)
            
            'Type3는 차액 열이 다름. 26열~28열
            est.차액 = .Cells(i, 26)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 27)
            est.수주금액 = .Cells(i, 28)
            
            'Type3는 수주차액이 없음
            est.수주차액 = ""
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 16 Or j > 20 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type3 = importCount
End Function

Function ImportEstimateData_Type4(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type4 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            est.입찰일 = Get_Date_Convert(.Cells(i, 12), year)
            est.수주일 = Get_Date_Convert(.Cells(i, 13), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 14), year)
            est.증권일 = Get_Date_Convert(.Cells(i, 15), year)
            
            '예상실행항목 등록
            'Type4는 Type3와 항목명 동일
            If .Cells(i, 16) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "외주(미르/대명/부성)"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주(현대기공/유진)"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "외주(근영/기타)"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.항목 = "명일(운송)"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 21)
            est.입찰금액 = .Cells(i, 22)
            
            'Type4는 차액 열이 다름
            est.차액 = .Cells(i, 26)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 27)
            
            'Type4는 수주금액, 수주차액이 없음
            est.수주금액 = ""
            est.수주차액 = ""
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 16 Or j > 20 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type4 = importCount
End Function

Function ImportEstimateData_Type5(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type5 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            est.입찰일 = Get_Date_Convert(.Cells(i, 12), year)
            est.수주일 = Get_Date_Convert(.Cells(i, 13), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 14), year)
            
            'Type5는 증권일이 없음
            est.증권일 = ""
            
            '예상실행항목 등록
            'Type5는 Type4와 항목명 동일
            If .Cells(i, 15) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 15)
                If WS.Cells(i, 15).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 15).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 16) <> "" Then
                prod.항목 = "외주(미르/대명/부성)"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "외주(현대기공/유진)"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주(근영/기타)"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "명일(운송)"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 20)
            est.입찰금액 = .Cells(i, 21)
            
            est.차액 = .Cells(i, 25)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 26)
            
            'Type5도 수주금액, 수주차액이 없음
            est.수주금액 = ""
            est.수주차액 = ""
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 15 Or j > 19 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type5 = importCount
End Function

Function ImportEstimateData_Type6(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type6 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 4 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            
            'Type6는 입찰일이 없음
            est.입찰일 = ""
            
            est.수주일 = Get_Date_Convert(.Cells(i, 12), year)
            est.납품일 = Get_Date_Convert(.Cells(i, 13), year)
            
            'Type6도 증권일이 없음
            est.증권일 = ""
            
            '예상실행항목 등록
            'Type6는 Type5와 항목명 동일
            If .Cells(i, 14) <> "" Then
                prod.항목 = "자재비"
                prod.비용 = .Cells(i, 14)
                If WS.Cells(i, 14).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 14).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 15) <> "" Then
                prod.항목 = "외주(미르/대명/부성)"
                prod.비용 = .Cells(i, 15)
                If WS.Cells(i, 15).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 15).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 16) <> "" Then
                prod.항목 = "외주(현대기공/유진)"
                prod.비용 = .Cells(i, 16)
                If WS.Cells(i, 16).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 16).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "외주(근영/기타)"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "명일(운송)"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            
            est.실행가 = .Cells(i, 19)
            est.입찰금액 = .Cells(i, 20)
            
            est.차액 = .Cells(i, 24)
            If est.입찰금액 = "" Or est.입찰금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 25)
            
            'Type6도 수주금액, 수주차액이 없음
            est.수주금액 = ""
            est.수주차액 = ""
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 14 Or j > 18 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type6 = importCount
End Function

Function ImportEstimateData_Type7(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type7 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(4, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 5 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            
            'Type7도 입찰일이 없음
            est.입찰일 = ""
            
            est.수주일 = Get_Date_Convert(.Cells(i, 12), year)
            
            'Type7의 납품일은 16번째 열에 있음
            est.납품일 = Get_Date_Convert(.Cells(i, 16), year)
            
            'Type7도 증권일이 없음
            est.증권일 = ""
            
            'Type7은 예상실행항목 2가지 있음
            If .Cells(i, 20) <> "" Then
                prod.항목 = "자재"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 21) <> "" Then
                prod.항목 = "외주"
                prod.비용 = .Cells(i, 21)
                If WS.Cells(i, 21).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 21).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, , prod.항목, , , , , prod.비용, prod.메모, regDate
            End If
            est.실행가 = ""
            
            'Type7의 입찰가는 수주금액으로 세팅함
            est.입찰금액 = ""
            est.차액 = ""
            
            'Type의 수주금액, 수주차액
            est.수주금액 = .Cells(i, 19)
            est.수주차액 = .Cells(i, 22)
            If est.수주금액 = "" Or est.수주금액 = "0" Then est.마진율 = "" Else est.마진율 = .Cells(i, 23)
            
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 20 Or j > 22 Then
                '예상실행항목 열이 아닌 경우만 실행
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type7 = importCount
End Function

Function ImportEstimateData_Type8(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As estimate
    Dim prod As Production
    Dim regDate As Date
    Dim pos As Integer
    Dim Y, M As String
    Dim currentId As Long
    
    Y = year
    pos = InStr(sheetName, "월")
    If pos <> 0 Then
        M = Left(sheetName, pos - 1)
        regDate = DateSerial(Y, M, 1)
    End If
    
    '1월~12월이 아니면 중지
    If M < 1 Or M > 12 Then
        ImportEstimateData_Type8 = 0
        Exit Function
    End If
    
    With WS
        endCol = .Cells(4, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        'Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
        'rng.Copy Sheet1.Cells(1, 1)
    End With
    
    importCount = 0
    For i = 5 To endRow
        currentId = Get_CurrentID(shtEstimateData)
        
        With WS
            est.관리번호 = .Cells(i, 1)
            est.자재번호 = .Cells(i, 2)
            est.거래처 = .Cells(i, 3)
            est.담당자 = .Cells(i, 4)
            est.품명 = .Cells(i, 5)
            
            '견적명이 없는 행은 문제있는 라인이므로 제외
            If est.품명 = "" Then
                GoTo NextIteration
            End If
            
            est.규격 = .Cells(i, 6)
            est.수량 = .Cells(i, 7)
            est.단위 = .Cells(i, 8)
            If .Cells(i, 7) <> "" Then      '#DIV/0! 오류 수정
                est.견적단가 = .Cells(i, 9)
            Else
                est.견적단가 = ""
            End If
            est.견적금액 = .Cells(i, 10)
            
            est.견적일 = Get_Date_Convert(.Cells(i, 11), year)
            
            'Type8도 입찰일이 없음
            est.입찰일 = ""
            
            est.수주일 = Get_Date_Convert(.Cells(i, 12), year)
            
            'Type8의 납품일은 16번째 열에 있음
            est.납품일 = Get_Date_Convert(.Cells(i, 16), year)
            
            'Type8도 증권일이 없음
            est.증권일 = ""
            
            'Type8은 예상실행항목 없음
            est.실행가 = ""
            
            'Type8의 입찰가는 수주금액으로 세팅함
            est.입찰금액 = ""
            est.차액 = ""
            est.마진율 = ""
            
            'Type8의 수주금액
            est.수주금액 = .Cells(i, 19)
            
            'Type8의 수주차액은 없음
            est.수주차액 = ""
            
            est.등록일자 = regDate
            
        End With
        
        '견적 등록
        Insert_Record shtEstimateData, est.관리번호, est.자재번호, est.거래처, est.담당자, est.품명, _
                est.규격, est.수량, est.단위, est.견적단가, est.견적금액, est.견적일, est.입찰일, est.수주일, est.납품일, est.증권일, _
                est.실행가, est.입찰금액, est.차액, est.마진율, est.수주금액, est.수주차액, est.등록일자, est.수정일자
        
        importCount = importCount + 1
        
        '메모 등록
        For j = 1 To endCol
            If j < 21 Then
                If Not WS.Cells(i, j).Comment Is Nothing Then
                    Insert_Record shtEstimateMemoData, currentId, est.관리번호, WS.Cells(i, j).Comment.Text, regDate
                End If
            End If
        Next
        
NextIteration:
    Next
    
    ImportEstimateData_Type8 = importCount
End Function

Function Get_Date_Convert(inputStr As String, Y As String)

    Dim pos As Long
    Dim M, D As String
    
    pos = InStr(inputStr, "/")
    If pos <> 0 Then
        M = Left(inputStr, pos - 1)
        D = Right(inputStr, Len(inputStr) - pos)
        If M = "" Or D = "" Then
            Get_Date_Convert = ""
        Else
            Get_Date_Convert = DateSerial(Y, M, D)
        End If
    Else
        Get_Date_Convert = inputStr
    End If
End Function
