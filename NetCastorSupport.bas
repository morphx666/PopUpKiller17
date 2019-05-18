Attribute VB_Name = "NetCastorsupport"
Option Explicit

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3

Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Function GetNCTabs(ByVal wH As Long) As String()

    Dim cName As String
    Dim Title As String
    Dim tabs() As String
    
    On Error GoTo ExitSub
    
    ReDim tabs(0)

    cName = Space(255)
    GetClassName wH, cName, 255
    If Left(cName, 8) = "TfrmMain" Then
        GetChildWindowHandles wH: wH = lcHandles(1)
        
        wH = GetNextWindow(wH, GW_HWNDNEXT)
        wH = GetNextWindow(wH, GW_HWNDNEXT)
        
        GetChildWindowHandles wH: wH = lcHandles(1)
        
        wH = GetNextWindow(wH, GW_HWNDNEXT)
        
        GetChildWindowHandles wH: wH = lcHandles(1)
        GetChildWindowHandles wH: wH = lcHandles(1)
        
        Do
            Title = Space(GetWindowTextLength(wH))
            GetWindowText wH, Title, 128
            
            ReDim Preserve tabs(UBound(tabs) + 1)
            tabs(UBound(tabs)) = Trim(Title) + Chr(0) & wH
            
            wH = GetNextWindow(wH, GW_HWNDNEXT)
        Loop While wH <> 0
    End If
    
ExitSub:
    
    GetNCTabs = tabs

End Function
