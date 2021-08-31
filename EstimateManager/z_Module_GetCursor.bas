Attribute VB_Name = "z_Module_GetCursor"
Option Explicit

'####################################################
'�� ���� ���콺 Ŀ���� ��ġ�� �ȼ��� ��ȯ�ϴ� ��ȯ�ϴ�User32 ������ API ����Դϴ�.
'    ���� �� ������ �����ο쳪, ���� �� ��ó�� �ݵ�� ����ؾ��մϴ�.
'    https://www.oppadu.com (�����ο���)
'-------------------------------------------------------
'�� �����
'Dim MousePOS As POINTAPI
'POINTAPI = convertMouseToForm
'������.Top = POINTAPI.Y
'������.Left = POINTAPI.X
'--------------------------------------------------------
'�� �� ����� �Ʒ� ��ũ�� �Ϻ� �����Ͽ� �ۼ��Ǿ����ϴ�.
'https://sites.google.com/a/mcpher.com/share/Home/excelquirks/snippets/mouseposition
'####################################################

#If VBA7 Then
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#Else
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Type POINTAPI
    x As Long
    y As Long
End Type

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Public Function pointsPerPixelX() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function pointsPerPixelY() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
End Function

Public Function WhereIsTheMouseAt() As POINTAPI
    Dim mPos As POINTAPI
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function

Public Function convertMouseToForm() As POINTAPI
    Dim mPos As POINTAPI
    mPos = WhereIsTheMouseAt
    mPos.x = pointsPerPixelY * mPos.x
    mPos.y = pointsPerPixelX * mPos.y
    convertMouseToForm = mPos
End Function
