Attribute VB_Name = "modFixMenuFocus"
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Sub SetMenuFocus(frm As Form)

    Dim h As Long

    h = GetSubMenu(GetMenu(frm.hwnd), 2)
    'Debug.Print h
    'Debug.Print SetActiveWindow(h)

End Sub
