VERSION 5.00
Begin VB.Form frmPUKCtrlIcon 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   4425
   ClientTop       =   4365
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   840
      Top             =   1380
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frmPUKCtrlIcon.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmPUKCtrlIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function apiSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub tmrRefresh_Timer()

    Dim wRect As RECT
    Dim actWin As Long
    Dim Title As String
    
    actWin = GetActiveWindow
    If actWin = hwnd Then
        'apiSetFocus GetDesktopWindow
    Else
        Title = Space(128)
        GetWindowText actWin, Title, 128
        If InStr(Title, "Microsoft Internet Explorer") > 0 Then Stop
        GetWindowRect actWin, wRect
        With wRect
            SetWindowPos hwnd, _
                        conHwndTopmost, _
                        .Right - 100, .Top + 5, _
                        16, 16, _
                        conSwpNoActivate Or conSwpShowWindow
        End With
    End If
    
End Sub
