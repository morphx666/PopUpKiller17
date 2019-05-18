Attribute VB_Name = "modAPI"
Option Explicit

Global Ok2Upload As Boolean
Global lHandles() As Long
Global lcHandles() As Long

Global HelpFile As String

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000

Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH As Long = &H101E

Public Const GWL_USERDATA = (-21)
Public Const GWL_STYLE = (-16)
Public Const GWL_ID = (-12)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_EXSTYLE = (-20)

Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_NOWAIT = &H2000

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const VK_CONTROL = &H11
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
   Const SE_ERR_NOASSOC = &H31
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub

Function EnumChildWindowsProc(ByVal wHandle As Long, lParam As Long) As Long

    ReDim Preserve lcHandles(UBound(lcHandles) + 1)
    lcHandles(UBound(lcHandles)) = wHandle
    EnumChildWindowsProc = True
    
End Function

Function EnumWindowsProc(ByVal wHandle As Long, lParam As Long) As Long

    If wTitle(wHandle) <> "" Then
        ReDim Preserve lHandles(UBound(lHandles) + 1)
        lHandles(UBound(lHandles)) = wHandle
    End If
    EnumWindowsProc = True
    
End Function

Sub GetWindowHandles()
    
    Dim lParam As Long
    ReDim lHandles(0)
    EnumWindows AddressOf EnumWindowsProc, lParam
    
End Sub

Sub GetChildWindowHandles(ParentID As Long)
    
    Dim lParam As Long
    ReDim lcHandles(0)
    EnumChildWindows ParentID, AddressOf EnumChildWindowsProc, lParam
    
End Sub

Public Sub AutoColumnsSize(lv As ListView)

    Dim i As Integer

    With lv
        For i = 0 To .ColumnHeaders.Count - 1
            SendMessageLong .hwnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER
        Next i
    End With
    
End Sub

Public Sub DisableCloseButton(frm As Form)

    Dim hMenu As Long
    Dim menuItemCount As Long
    
    'Obtain the handle to the form's system menu
    hMenu = GetSystemMenu(frm.hwnd, 0)
    
    If hMenu Then
     
        'Obtain the number of items in the menu
         menuItemCount = GetMenuItemCount(hMenu)
        
        'Remove the system menu Close menu item.
        'The menu item is 0-based, so the last
        'item on the menu is menuItemCount - 1
         Call RemoveMenu(hMenu, menuItemCount - 1, MF_REMOVE Or MF_BYPOSITION)
        
        'Remove the system menu separator line
         Call RemoveMenu(hMenu, menuItemCount - 2, MF_REMOVE Or MF_BYPOSITION)
        
        'Force a redraw of the menu. This
        'refreshes the titlebar, dimming the X
         Call DrawMenuBar(frm.hwnd)
    
    End If

End Sub

Public Function GetDlgText(hwnd As Long) As String

    Dim ii As Integer
    Dim Title As String

    GetChildWindowHandles hwnd
    For ii = 1 To UBound(lcHandles)
        If clsName(lcHandles(ii)) = "Static" Then
            Title = Space(SendMessageLong(lcHandles(ii), WM_GETTEXTLENGTH, 0, ByVal 0))
            SendMessageLong lcHandles(ii), WM_GETTEXT, Len(Title) + 1, ByVal Title
            If Title <> "" Then
                GetDlgText = Title
                Exit Function
            End If
        End If
    Next ii

End Function

Public Function wTitle(wHnd As Long) As String

    wTitle = Space(GetWindowTextLength(wHnd))
    GetWindowText wHnd, wTitle, Len(wTitle) + 1

End Function

Public Function clsName(wHnd As Long) As String

    On Error Resume Next

    clsName = Space(255)
    GetClassName wHnd, clsName, 255
    clsName = Left(clsName, InStr(clsName, Chr(0)) - 1)

End Function
