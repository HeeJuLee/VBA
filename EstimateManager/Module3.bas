Attribute VB_Name = "Module3"
Option Explicit

Sub 매크로5()
Attribute 매크로5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로5 매크로
'

'
    Range("D10:AC10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("I16").Select
End Sub
Sub 매크로6()
Attribute 매크로6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로6 매크로
'

'
    Range("I8").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub
Sub 매크로7()
Attribute 매크로7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로7 매크로
'

'
    Range("H10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
End Sub
Sub 매크로8()
Attribute 매크로8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로8 매크로
'

'
    Range("N5:S19").Select
    Selection.AutoFilter
    Range("P10").Select
    Selection.AutoFilter
End Sub
Sub 매크로9()
Attribute 매크로9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로9 매크로
'

'
    Range("N5:S17").Select
    Selection.AutoFilter
    
    Range("P11").Select
    Selection.AutoFilter
End Sub
