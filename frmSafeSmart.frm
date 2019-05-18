VERSION 5.00
Begin VB.Form frmSafeSmart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Safe Smart!"
   ClientHeight    =   2625
   ClientLeft      =   5175
   ClientTop       =   5700
   ClientWidth     =   5940
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   15
      ScaleHeight     =   1650
      ScaleWidth      =   2505
      TabIndex        =   5
      Top             =   915
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddBlackList 
      Caption         =   "Add to Black List"
      Height          =   345
      Left            =   4335
      TabIndex        =   4
      Top             =   2235
      Width           =   1605
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close It"
      Height          =   345
      Left            =   4335
      TabIndex        =   3
      Top             =   1830
      Width           =   1605
   End
   Begin VB.CommandButton cmdAddIgnore 
      Caption         =   "Add to Ignore List"
      Height          =   345
      Left            =   2625
      TabIndex        =   2
      Top             =   2235
      Width           =   1605
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore It"
      Height          =   345
      Left            =   2625
      TabIndex        =   1
      Top             =   1815
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmSafeSmart.frx":0000
      Stretch         =   -1  'True
      Top             =   262
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "The Smart! engine..."
      Height          =   795
      Left            =   1050
      TabIndex        =   0
      Top             =   105
      Width           =   4815
   End
End
Attribute VB_Name = "frmSafeSmart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40


Private Sub cmdAddBlackList_Click()

    SafeSmartAns = CloseIt_and_Add2BlackList
    Unload Me

End Sub

Private Sub cmdAddIgnore_Click()

    SafeSmartAns = IgnoreIt_and_Add2IgnoreList
    Unload Me

End Sub

Private Sub cmdClose_Click()

    SafeSmartAns = CloseIt
    Unload Me

End Sub

Private Sub cmdIgnore_Click()

    SafeSmartAns = IgnoreIt
    Unload Me

End Sub

Private Sub Form_Load()

    Dim cForeground As Long
    Dim hDC As Long
    
    SetupCharset Me

    cForeground = GetForegroundWindow()
    SetForegroundWindow SafeSmartHwnd
    SetFocusAPI SafeSmartHwnd
    
    hDC = GetDC(SafeSmartHwnd)

    BitBlt picSS.hDC, 0, 0, picSS.ScaleWidth, picSS.ScaleHeight, _
           hDC, 0, 0, vbSrcCopy
           
    DoEvents
    
    Width = picSS.Width + cmdIgnore.Width * 2 + GetClientLeft(Me.hwnd) + 240
    Height = cmdAddIgnore.Top + cmdAddIgnore.Height + GetClientTop(Me.hwnd) + 60

    Move Screen.Width / 2 - Width / 2, Screen.Height / 2 - Height / 2
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
                Left / Screen.TwipsPerPixelX, Top / Screen.TwipsPerPixelY, _
                Width / Screen.TwipsPerPixelX, Height / Screen.TwipsPerPixelY, _
                SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    BringWindowToTop cForeground
    DoEvents
    SetForegroundWindow cForeground
    SetFocusAPI cForeground

End Sub
