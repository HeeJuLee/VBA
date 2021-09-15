Attribute VB_Name = "Module1"
Option Explicit

Sub 매크로1()
Attribute 매크로1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로1 매크로
'

'
    ActiveSheet.Unprotect
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowFiltering:=True
End Sub
