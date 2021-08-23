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
Type Estimate
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
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ClearEstimateData
    
'    Set WB = Application.Workbooks.Open("C:\Users\leehe\Downloads\현솔-견적관리문서\법인-견적관리2013.xlsx")
'    For Each WS In WB.Worksheets
'        If WS.Name Like "**월" Then
'            importCount = ImportEstimateData_Type1(WS, WS.Name, 2013)
'        End If
'    Next
'    WB.Close
    
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
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtProductionData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
    With shtEstimateMemoData
        endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        endRow = .Cells(.Rows.Count, 1).End(xlUp).row
        .Cells(1, endCol) = 100000
        .Range("A2").Resize(endRow, endCol).Delete
    End With
    
End Sub

Function ImportEstimateData_Type1(WS As Worksheet, sheetName As String, year As String) As Long

    Dim i, j As Long
    Dim endCol As Long
    Dim endRow As Long
    Dim importCount As Long
    Dim est As Estimate
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
                Insert_Record shtProductionData, currentId, est.관리번호, prod.항목, prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 17) <> "" Then
                prod.항목 = "미르"
                prod.비용 = .Cells(i, 17)
                If WS.Cells(i, 17).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 17).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, prod.항목, prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 18) <> "" Then
                prod.항목 = "외주"
                prod.비용 = .Cells(i, 18)
                If WS.Cells(i, 18).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 18).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, prod.항목, prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 19) <> "" Then
                prod.항목 = "인건비"
                prod.비용 = .Cells(i, 19)
                If WS.Cells(i, 19).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 19).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, prod.항목, prod.비용, prod.메모, regDate
            End If
            If .Cells(i, 20) <> "" Then
                prod.항목 = "기타"
                prod.비용 = .Cells(i, 20)
                If WS.Cells(i, 20).Comment Is Nothing Then prod.메모 = "" Else prod.메모 = WS.Cells(i, 20).Comment.Text
                Insert_Record shtProductionData, currentId, est.관리번호, prod.항목, prod.비용, prod.메모, regDate
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

Function Get_Date_Convert(inputStr As String, Y As String)

    Dim pos As Long
    Dim M, D As String
    
    pos = InStr(inputStr, "/")
    If pos <> 0 Then
        M = Left(inputStr, pos - 1)
        D = Right(inputStr, Len(inputStr) - pos)
        Get_Date_Convert = DateSerial(Y, M, D)
    Else
        Get_Date_Convert = inputStr
    End If
End Function
