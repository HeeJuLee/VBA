Attribute VB_Name = "Mod_WS"
Option Explicit

Public clickOrderId As Variant
Public estimateUpdateFormX, estimateUpdateFormY As Long
Public orderUpdateFormX, orderUpdateFormY As Long
Public estimateInsertFormX, estimateInsertFormY As Long
Public orderInsertFormX, orderInsertFormY As Long
Public selectionRow As Long

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

Sub ClearContentsLine(startRng As Range, endColNo, clearRowCount)
    
    Dim WS As Worksheet
    Dim lastRow As Long
    Set WS = startRng.Parent
        
    If Not IsNumeric(endColNo) Then
        endColNo = Range(endColNo & 1).Column
    End If
        
    lastRow = startRng.row + clearRowCount
    If lastRow < startRng.row Then Exit Sub
    
    With WS.Range(startRng, WS.Cells(lastRow, endColNo))
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

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

