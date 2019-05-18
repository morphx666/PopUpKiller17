VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   8595
   ClientLeft      =   4500
   ClientTop       =   3660
   ClientWidth     =   12000
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
   ScaleHeight     =   8595
   ScaleWidth      =   12000
   Begin VB.Frame frameDetection 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   6975
      TabIndex        =   18
      Top             =   2790
      Width           =   4740
      Begin VB.CheckBox chkDisWildCards 
         Caption         =   "Disable Wildcards in Black List"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   4155
         Width           =   4230
      End
      Begin VB.CheckBox chkFGeo 
         Caption         =   "GeoCities BOX Support"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   4410
         Width           =   4230
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   285
         Left            =   3510
         TabIndex        =   15
         Top             =   3075
         Width           =   735
      End
      Begin VB.CheckBox chkSafeMode 
         Caption         =   "Safe Mode"
         Height          =   210
         Left            =   1725
         TabIndex        =   10
         Top             =   570
         Width           =   2490
      End
      Begin VB.TextBox txtCustomIETitle 
         Height          =   285
         Left            =   390
         TabIndex        =   14
         Text            =   "Microsoft Internet Explorer"
         Top             =   3075
         Width           =   3075
      End
      Begin VB.CheckBox chkCustomTitle 
         Caption         =   "Custom Internet Explorer Title"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   2850
         Width           =   4230
      End
      Begin VB.TextBox txtLimitNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Text            =   "33"
         Top             =   1980
         Width           =   315
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit Simultaneous Windows"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   2010
         Width           =   2535
      End
      Begin VB.CheckBox chkSmart 
         Caption         =   "Smart!"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   4230
      End
      Begin MSComctlLib.Slider sldSmart 
         Height          =   225
         Left            =   45
         TabIndex        =   9
         Top             =   555
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   1
         Min             =   15
         Max             =   35
         SelStart        =   15
         Value           =   15
         TextPosition    =   1
      End
      Begin VB.Label lblCustomIETitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":0000
         ForeColor       =   &H00808080&
         Height          =   600
         Left            =   120
         TabIndex        =   21
         Top             =   3390
         Width           =   4590
      End
      Begin VB.Label lblLimitMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Enable this option to control the maximum number of browser windows to be opened at the same time. "
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   2250
         Width           =   4590
      End
      Begin VB.Label lblSmart 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":008E
         ForeColor       =   &H00808080&
         Height          =   1020
         Left            =   120
         TabIndex        =   19
         Top             =   825
         Width           =   4695
      End
   End
   Begin VB.Frame frameLookFeel 
      BorderStyle     =   0  'None
      Caption         =   "Look && Feel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   5580
      TabIndex        =   23
      Top             =   780
      Width           =   4740
      Begin VB.CheckBox chkSound 
         Caption         =   "Enable Sound Effects"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3915
         Width           =   3960
      End
      Begin VB.CheckBox chkScrambleText 
         Caption         =   "Scramble Popups Text"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2940
         Width           =   3960
      End
      Begin VB.CheckBox chkShowTabsText 
         Caption         =   "Show Tabs Text"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1995
         Width           =   3960
      End
      Begin VB.CheckBox chkShowIcon 
         Caption         =   "Show Tray Icon"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   1035
         Width           =   3960
      End
      Begin VB.CheckBox chkFlashIcon 
         Caption         =   "Flash Tray Icon"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   3960
      End
      Begin VB.Label lblScrambleText 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":01A4
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   3165
         Width           =   4365
      End
      Begin VB.Label lblShowTabsText 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":024C
         ForeColor       =   &H00808080&
         Height          =   585
         Left            =   120
         TabIndex        =   26
         Top             =   2235
         Width           =   4695
      End
      Begin VB.Label lblShowIcon 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":02DF
         ForeColor       =   &H00808080&
         Height          =   585
         Left            =   120
         TabIndex        =   25
         Top             =   1290
         Width           =   4695
      End
      Begin VB.Label lblFlashIcon 
         BackStyle       =   0  'Transparent
         Caption         =   "If you enable this option, PUK will flash its icon on the tray bar every time a popup is killed."
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   570
         Width           =   4695
      End
   End
   Begin VB.Frame frameGeneral 
      BorderStyle     =   0  'None
      Caption         =   "General"
      Height          =   5190
      Left            =   120
      TabIndex        =   28
      Top             =   390
      Width           =   4740
      Begin VB.CheckBox chkDisSCKeys 
         Caption         =   "Disable Shortcut Keys"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2460
         Width           =   4350
      End
      Begin VB.TextBox txtHomePage 
         Height          =   285
         Left            =   390
         TabIndex        =   2
         Text            =   "http://software.xfx.net/"
         Top             =   1365
         Width           =   3540
      End
      Begin VB.CheckBox chkAutoJump 
         Caption         =   "Jump to Home Page"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   4350
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "Auto Start"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   4350
      End
      Begin VB.Label lblDisSCKeys 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Check this option to avoid PopUp Killer from going into idle mode when pressing the CTRL key or the CTRL+SHIFT keys combination."
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   120
         TabIndex        =   34
         Top             =   2700
         Width           =   4695
      End
      Begin VB.Label lblAutoJump 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Enable this option to make PUK open a new browser window and jump to your home page in case PUK closes all the opened windows."
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lblAutoStart 
         BackStyle       =   0  'Transparent
         Caption         =   "Enable this option to have PUK start automatically every time Windows starts."
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   120
         TabIndex        =   29
         Top             =   540
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4170
      TabIndex        =   17
      Top             =   5685
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3225
      TabIndex        =   16
      Top             =   5685
      Width           =   825
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   5565
      Left            =   15
      TabIndex        =   22
      Top             =   45
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   9816
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "tsGeneral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Look && Feel"
            Key             =   "tsLookFeel"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detection"
            Key             =   "tsDetection"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ResetColors(Optional xLBL As Label)

    Dim ctrl As Object
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is Label Then
            If xLBL Is Nothing Then
                ctrl.ForeColor = &H808080
            Else
                If ctrl.Name <> xLBL.Name Then ctrl.ForeColor = &H808080
            End If
        End If
    Next ctrl

End Sub

Private Sub chkAutoJump_Click()

    UpdateControls

End Sub

Private Sub chkAutoJump_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblAutoJump
    lblAutoJump.ForeColor = vbBlack

End Sub

Private Sub chkAutoStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblAutoStart
    lblAutoStart.ForeColor = vbBlack

End Sub

Private Sub chkCustomTitle_Click()

    UpdateControls

End Sub

Private Sub chkCustomTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblCustomIETitle
    lblCustomIETitle.ForeColor = vbBlack

End Sub

Private Sub chkDisSCKeys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblDisSCKeys
    lblDisSCKeys.ForeColor = vbBlack

End Sub

Private Sub chkFlashIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblFlashIcon
    lblFlashIcon.ForeColor = vbBlack

End Sub

Private Sub chkLimit_Click()

    UpdateControls

End Sub

Private Sub chkLimit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblLimitMode
    lblLimitMode.ForeColor = vbBlack

End Sub

Private Sub chkSafeMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblSmart
    lblSmart.ForeColor = vbBlack

End Sub

Private Sub chkScrambleText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblScrambleText
    lblScrambleText.ForeColor = vbBlack

End Sub

Private Sub chkShowIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblShowIcon
    lblShowIcon.ForeColor = vbBlack

End Sub

Private Sub chkShowTabsText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblShowTabsText
    lblShowTabsText.ForeColor = vbBlack

End Sub

Private Sub chkSmart_Click()

    UpdateControls

End Sub

Private Sub chkSmart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblSmart
    lblSmart.ForeColor = vbBlack

End Sub

Private Sub cmdCancel_Click()

    Preferences.Changed = False
    Unload Me

End Sub

Private Sub cmdDefault_Click()

    Dim DetectedTitle As String
    
    DetectedTitle = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title")
    If DetectedTitle = "" Then DetectedTitle = "Microsoft Internet Explorer"

    txtCustomIETitle.Text = DetectedTitle
    txtCustomIETitle.SetFocus

End Sub

Private Sub cmdDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblCustomIETitle
    lblCustomIETitle.ForeColor = vbBlack

End Sub

Private Sub cmdOK_Click()

    With Preferences
        .AutoStart = -chkAutoStart.Value
        .JumpToHomePage = -chkAutoJump.Value
        .HomePage = txtHomePage.Text
        
        .FlashIcon = -chkFlashIcon.Value
        .ShowIcon = -chkShowIcon.Value
        .ShowTabsText = -chkShowTabsText.Value
        .ScrambleText = -chkScrambleText.Value
        .SoundFX = -chkSound.Value
        
        .Smart = -chkSmart.Value
        .SafeMode = -chkSafeMode.Value
        .SmartSensitivity = sldSmart.Value * 10000
        .Limit = -chkLimit.Value
        .LimitNum = txtLimitNum.Text
        .UseCustomIETitle = -chkCustomTitle.Value
        .CustomIETitle = txtCustomIETitle.Text
        
        If .UseCustomIETitle And .CustomIETitle = "" Then
            .UseCustomIETitle = False
            MsgBox "The ""Custom Internet Explorer Title"" field cannot be left empty", vbInformation + vbOKOnly, "Invalid Parameters"
            Exit Sub
        End If
        
        .FGeo = -chkFGeo.Value
        
        .DisableShortcutKeys = -chkDisSCKeys.Value
        .DisableWildcards = -chkDisWildCards.Value
        
        .Changed = True
    End With
    
    Unload Me

End Sub

Private Sub Form_Load()

    Width = 5145
    Height = 6075 + GetClientTop(Me.hwnd)
    
    SkinMe Me, True
    SetupCharset Me
    
    With Screen
        Move .Width / 2 - Width / 2, .Height / 2 - Height / 2
    End With
    
    With frameGeneral
        frameDetection.Move .Left, .Top
        frameLookFeel.Move .Left, .Top
        .ZOrder 0
    End With
    
    With Preferences
        chkAutoStart.Value = Abs(.AutoStart)
        chkAutoJump.Value = Abs(.JumpToHomePage)
        txtHomePage.Text = .HomePage
        
        chkFlashIcon.Value = Abs(.FlashIcon)
        chkShowIcon.Value = Abs(.ShowIcon)
        chkShowTabsText.Value = Abs(.ShowTabsText)
        chkScrambleText.Value = Abs(.ScrambleText)
        chkSound.Value = Abs(.SoundFX)
        
        chkSmart.Value = Abs(.Smart)
        chkSafeMode.Value = Abs(.SafeMode)
        sldSmart.Value = .SmartSensitivity / 10000
        chkLimit.Value = Abs(.Limit)
        txtLimitNum.Text = .LimitNum
        txtLimitNum.Enabled = .Limit
        chkCustomTitle.Value = Abs(.UseCustomIETitle)
        txtCustomIETitle.Text = .CustomIETitle
        chkFGeo.Value = Abs(.FGeo)
        chkFGeo.Enabled = Not (FGeoLib Is Nothing)
        
        chkDisSCKeys.Value = Abs(.DisableShortcutKeys)
        chkDisWildCards.Value = Abs(.DisableWildcards)
    End With
    
    UpdateControls

End Sub

Private Sub UpdateControls()

    txtCustomIETitle.Enabled = -chkCustomTitle.Value
    cmdDefault.Enabled = -chkCustomTitle.Value
    txtHomePage.Enabled = -chkAutoJump.Value
    chkSafeMode.Enabled = -chkSmart.Value
    sldSmart.Enabled = -chkSmart.Value
    txtLimitNum.Enabled = -chkLimit.Value

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SkinMe Me, False

End Sub

Private Sub frameDetection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors

End Sub

Private Sub frameGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors

End Sub

Private Sub frameLookFeel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors

End Sub

Private Sub sldSmart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblSmart
    lblSmart.ForeColor = vbBlack

End Sub

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.Key
        Case "tsGeneral"
            frameGeneral.ZOrder 0
        Case "tsLookFeel"
            frameLookFeel.ZOrder 0
        Case "tsDetection"
            frameDetection.ZOrder 0
    End Select

End Sub

Private Sub txtCustomIETitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblCustomIETitle
    lblCustomIETitle.ForeColor = vbBlack

End Sub

Private Sub txtHomePage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblAutoJump
    lblAutoJump.ForeColor = vbBlack

End Sub

Private Sub txtLimitNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors lblLimitMode
    lblLimitMode.ForeColor = vbBlack

End Sub
