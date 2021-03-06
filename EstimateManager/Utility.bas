Attribute VB_Name = "Utility"
Option Explicit

Public clickOrderId, currentOrderId As Variant
Public doubleClickFlag, clickEstimateId, currentEstimateId, currentManagementId, currentAcceptedId As Variant
Public estimateUpdateFormX, estimateUpdateFormY As Long
Public orderUpdateFormX, orderUpdateFormY As Long
Public estimateInsertFormX, estimateInsertFormY As Long
Public orderInsertFormX, orderInsertFormY As Long
Public productionFormX, productionFormY As Long
Public paymentFormX, paymentFormY As Long
Public operationFormX, operationFormY As Long
Public orderDateUpdateFormX, orderDateUpdateFormY As Long
Public productionCopyFormX, productionCopyFormY As Long
Public selectionRow As Long
Public bDeleteFlag As Boolean

Sub GetCalendarDate(textBox As MSForms.textBox)
    Dim vDate As Date
    Dim orgValue As Variant
    
    orgValue = textBox.value
    
    vDate = frmCalendar.GetDate
    
    ' X나 ESC를 눌러서 나온 경우, 날짜 포맷이 '오전 10:00' 이런식으로 넘어옴. 오류 체크
    If InStr(vDate, "오전") <> 0 Or InStr(vDate, "오후") <> 0 Then
        'X를 누른 경우에 미리 입력되어 있던 값이 있으면 유지
        If orgValue = "" Then
            textBox.value = ""
        End If
    Else
        textBox.value = vDate
    End If
End Sub

Function GetCalendarDate_2(orgValue)
    Dim vDate As Date
    
    vDate = frmCalendar.GetDate
    
    ' X나 ESC를 눌러서 나온 경우, 날짜 포맷이 '오전 10:00' 이런식으로 넘어옴. 오류 체크
    If InStr(vDate, "오전") <> 0 Or InStr(vDate, "오후") <> 0 Then
        'X를 누른 경우에 미리 입력되어 있던 값이 있으면 유지
        If orgValue = "" Then
            GetCalendarDate = ""
        Else
            GetCalendarDate = orgValue
        End If
    Else
        GetCalendarDate = vDate
    End If
End Function

Public Function getLocalFullName$(ByVal fullPath$)
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02

    Dim ii&
    Dim iPos&
    Dim oneDrivePath$
    Dim endFilePath$

    If Left(fullPath, 8) = "https://" Then 'Possibly a OneDrive URL
        If InStr(1, fullPath, "my.sharepoint.com") <> 0 Or InStr(1, fullPath, "https://onedrive.") <> 0 Then 'Commercial OneDrive
            'For commercial OneDrive, path looks like
            ' "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            'Find "/Documents" in string and replace everything before the end with OneDrive local path
            iPos = InStr(1, fullPath, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
            endFilePath = mid(fullPath, iPos) 'Get the ending file path without pointer in OneDrive. Include leading "/"
        Else 'Personal OneDrive
            'For personal OneDrive, path looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
            'We can get local file path by replacing "https.." up to the 4th slash, with the OneDrive local path obtained from registry
            iPos = 8 'Last slash in https://
            For ii = 1 To 2
                iPos = InStr(iPos + 1, fullPath, "/") 'find 4th slash
            Next ii
            endFilePath = mid(fullPath, iPos) 'Get the ending file path without OneDrive root. Include leading "/"
        End If
        endFilePath = Replace(endFilePath, "/", Application.PathSeparator) 'Replace forward slashes with back slashes (URL type to Windows type)
        
        'getLocalFullName = getLocalOneDrivePath & endFilePath
        
        For ii = 1 To 3 'Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists, if so return the name
            oneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive")) 'Check possible local paths. "OneDrive" should be the last one
            If 0 < Len(oneDrivePath) Then
                getLocalFullName = oneDrivePath & endFilePath
                Exit Function 'Success (i.e. found the correct Environ parameter)
            End If
        Next ii
        'Possibly raise an error here when attempt to convert to a local file name fails - e.g. for "shared with me" files
        getLocalFullName = vbNullString
    Else
        getLocalFullName = fullPath
    End If
End Function

Sub MoveToEstimateAdmin()
    
    shtEstimateAdmin.Activate
    
End Sub

Sub MoveToOrderAdmin()
    
    shtOrderAdmin.Activate
    
End Sub

Sub MoveToOperationAdmin()
    
    shtOperationAdmin.Activate
    
End Sub

Sub MoveToFinance()
    
    shtFinance.Activate
    
End Sub

Sub SetContentsLine(startRng As Range, endColNo, clearRowCount)
    Dim WS As Worksheet
    Dim lastRow As Long
    Set WS = startRng.Parent
        
    If Not IsNumeric(endColNo) Then
        endColNo = Range(endColNo & 1).Column
    End If
        
    lastRow = startRng.row + clearRowCount - 1
    If lastRow < startRng.row Then Exit Sub
    
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
End Sub

'Sub ClearContentsLine(startRng As Range, endColNo, clearRowCount)
Sub ClearContentsLine(startRng As Range, endColNo)
    
    Dim WS As Worksheet
    Dim lastRow As Long
    Set WS = startRng.Parent
        
    If Not IsNumeric(endColNo) Then
        endColNo = Range(endColNo & 1).Column
    End If
        
    'lastRow = startRng.row + clearRowCount
    lastRow = startRng.End(xlDown).row
    If lastRow < startRng.row Then Exit Sub
    
    '라인 서식 지우기
    With WS.Range(startRng, WS.Cells(lastRow, endColNo))
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    '채우기 색상 지우기
    With WS.Range(startRng, WS.Cells(lastRow, endColNo)).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    '메모 지우기
    With WS.Range(startRng, WS.Cells(lastRow, endColNo))
        .ClearComments
    End With

End Sub

Sub SetContentsColor(startRng As Range, endColNo, arr, colNo, strMatch, color)
    Dim WS As Worksheet
    Dim currentRow, startColNo As Long
    Dim i As Long
    Set WS = startRng.Parent
        
    currentRow = startRng.row
    startColNo = startRng.Column
    
    If Not IsNumeric(endColNo) Then
        endColNo = Range(endColNo & 1).Column
    End If
    
    For i = 1 To UBound(arr)
        If arr(i, colNo) = strMatch Then
             With WS.Range(WS.Cells(currentRow, startColNo), WS.Cells(currentRow, endColNo)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                '.ThemeColor = color
                .color = color
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End If
        currentRow = currentRow + 1
    Next
End Sub

Sub SetReceivingColor(startRng As Range, arr, orderColNo, reveivingColNo, color)
    Dim WS As Worksheet
    Dim currentRow, currentColNo As Long
    Dim i As Long
    Set WS = startRng.Parent
        
    currentRow = startRng.row
    currentColNo = startRng.Column
    
    For i = 1 To UBound(arr)
        If arr(i, orderColNo) <> "" And arr(i, reveivingColNo) = "" Then
             With WS.Cells(currentRow, currentColNo).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                '.ThemeColor = color
                .color = color
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End If
        currentRow = currentRow + 1
    Next
End Sub

Sub SetComment(db, memoColNo As Long, startRng As Range)
    Dim WS As Worksheet
    Dim i As Long
    Dim currentRow As Long
    Set WS = startRng.Parent
    
    currentRow = startRng.row
    
    For i = 1 To UBound(db)
        If db(i, memoColNo) <> "" Then
            WS.Cells(currentRow, startRng.Column).AddComment db(i, memoColNo)
            WS.Cells(currentRow, startRng.Column).Comment.Shape.ScaleWidth 1.5, msoFalse
            If Len(db(i, memoColNo)) > 50 Then
                WS.Cells(currentRow, startRng.Column).Comment.Shape.ScaleHeight 1.5, msoFalse
            End If
        End If
        currentRow = currentRow + 1
    Next

End Sub

Sub SetHyperLink(db, itemColNo As Long, rng As Range)
    Dim WS As Worksheet
    Dim i As Long
    Dim currentRow As Long
    Set WS = rng.Parent
    
    currentRow = rng.row
    
    For i = 1 To UBound(db)
        If db(i, itemColNo) = "" Then
            db(i, itemColNo) = "(없음)"
        End If
            
        WS.Hyperlinks.Add WS.Cells(currentRow, rng.Column), "", "", , db(i, itemColNo)
        
        currentRow = currentRow + 1
    Next

End Sub

Sub SetHyperLink2(db, manIdColNo As Long, itemColNo As Long, rng As Range, rng2 As Range)
    Dim WS As Worksheet
    Dim i As Long
    Dim currentRow As Long
    Set WS = rng.Parent
    
    currentRow = rng.row
    
    For i = 1 To UBound(db)
        If db(i, itemColNo) = "" Then
            db(i, itemColNo) = "(없음)"
        End If
            
        WS.Hyperlinks.Add WS.Cells(currentRow, rng.Column), "", "", , db(i, manIdColNo)
        WS.Hyperlinks.Add WS.Cells(currentRow, rng2.Column), "", "", , db(i, itemColNo)
        
        currentRow = currentRow + 1
    Next
End Sub

Function isFormLoaded(ByVal strName As String) As Boolean
    Dim i As Integer

    isFormLoaded = True
    strName = LCase(strName)
    For i = 0 To VBA.UserForms.count - 1
        If LCase(UserForms(i).Name) = strName Then Exit Function
    Next
    isFormLoaded = False
End Function


'==========================================================================================
Private Sub InitializeCboCustomer()
    Dim db As Variant
    db = Get_DB(shtEstimateCustomer, True)

    Update_Cbo Me.cboCustomer, db, 1
End Sub

Private Sub InitializeCboManager()
    Dim db As Variant
    
    '담당자 DB를 읽어와서
    db = Get_DB(shtEstimateManager, True)
    '거래처명으로 필터링
    db = Filtered_DB(db, Me.cboCustomer.value, 1, True)
    
    '기존 콤보박스 내용지우기
    Me.cboManager.Clear
    
    '담당자가 있으면 콤보박스에 추가함
    If Not isEmpty(db) Then
        Update_Cbo Me.cboManager, db, 2
    End If
End Sub

Function ConvertDateFormat(value)
    Dim pos As Long
    Dim Y, M, D As Long
    
    ConvertDateFormat = ""
    
    If value = "" Then
        Exit Function
    End If
    
    pos = InStr(value, "월")
    If pos > 0 Then
        M = Left(value, pos - 1)
        If M <> "" And IsNumeric(M) Then
            ConvertDateFormat = DateSerial(Year(Date), M, 1)
            Exit Function
        End If
    End If
    
    pos = InStr(value, "/")
    If pos > 0 Then
        M = Left(value, pos - 1)
        If M = "" Then
            M = month(Date)
        End If
        If Len(value) = pos Then
            D = 1
        ElseIf IsNumeric(mid(value, pos + 1)) Then
            D = mid(value, pos + 1)
        End If
        ConvertDateFormat = DateSerial(Year(Date), M, D)
    ElseIf Len(value) = 4 And IsNumeric(value) Then
        '4자리 숫자
        M = Left(value, 2)
        D = Right(value, 2)
        ConvertDateFormat = DateSerial(Year(Date), M, D)
    ElseIf Len(value) <= 2 And IsNumeric(value) Then
        '1자리/2자리 숫자
        M = value
        ConvertDateFormat = DateSerial(Year(Date), M, 1)
    Else
        If IsDate(value) = True Then
            ConvertDateFormat = Format(value, "yyyy-mm-dd")
        End If
    End If
    
End Function

Function isExistInSheet(startRng As Range, value) As Long
    Dim WS As Worksheet
    Dim lastRow As Long
    Dim col As Long
    Dim i As Long
    Set WS = startRng.Parent
    
    lastRow = startRng.End(xlDown).row
    col = startRng.Column
    
    If IsNumeric(value) Then
        value = CLng(value)
    End If
    
    isExistInSheet = 0
    For i = startRng.row To lastRow
        If WS.Cells(i, col) = value Then
            isExistInSheet = i
            Exit Function
        End If
    Next
End Function

Sub UndoLastAction()

    With Application
        .EnableEvents = False
        .Undo
        .EnableEvents = True
    End With

End Sub
