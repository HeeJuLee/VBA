Attribute VB_Name = "Mod_WS"
Option Explicit

Public clickOrderId, currentOrderId As Variant
Public clickEstimateId, currentEstimateId As Variant
Public estimateUpdateFormX, estimateUpdateFormY As Long
Public orderUpdateFormX, orderUpdateFormY As Long
Public estimateInsertFormX, estimateInsertFormY As Long
Public orderInsertFormX, orderInsertFormY As Long
Public productionFormX, productionFormY As Long
Public paymentFormX, paymentFormY As Long
Public operationFormX, operationFormY As Long
Public orderDateUpdateFormX, orderDateUpdateFormY As Long
Public selectionRow As Long

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
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without pointer in OneDrive. Include leading "/"
        Else 'Personal OneDrive
            'For personal OneDrive, path looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
            'We can get local file path by replacing "https.." up to the 4th slash, with the OneDrive local path obtained from registry
            iPos = 8 'Last slash in https://
            For ii = 1 To 2
                iPos = InStr(iPos + 1, fullPath, "/") 'find 4th slash
            Next ii
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without OneDrive root. Include leading "/"
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


Sub SetComment(db, memoColNo As Long, startRng As Range)
    Dim WS As Worksheet
    Dim i As Long
    Dim currentRow As Long
    Set WS = startRng.Parent
    
    currentRow = startRng.row
    
    For i = 1 To UBound(db)
        If db(i, memoColNo) <> "" Then
            WS.Cells(currentRow, startRng.Column).AddComment db(i, memoColNo)
        End If
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
    If Not IsEmpty(db) Then
        Update_Cbo Me.cboManager, db, 2
    End If
End Sub
