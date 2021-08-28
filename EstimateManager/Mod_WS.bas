Attribute VB_Name = "Mod_WS"
Option Explicit


Sub GetCalendarDate(textBox As MSForms.textBox)
    Dim vDate As Date
    Dim orgValue As Variant
    
    orgValue = textBox.Value
    
    vDate = frmCalendar.GetDate
    
    ' X나 ESC를 눌러서 나온 경우, 날짜 포맷이 '오전 10:00' 이런식으로 넘어옴. 오류 체크
    If InStr(vDate, "오전") <> 0 Or InStr(vDate, "오후") <> 0 Then
        'X를 누른 경우에 미리 입력되어 있던 값이 있으면 유지
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
