VERSION 5.00
Begin VB.Form frmDetModeWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Learn Mode Wizard"
   ClientHeight    =   6345
   ClientLeft      =   2550
   ClientTop       =   3345
   ClientWidth     =   10395
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
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
   ScaleHeight     =   6345
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   3705
      ScaleHeight     =   2625
      ScaleWidth      =   3030
      TabIndex        =   15
      Tag             =   "3"
      Top             =   3585
      Width           =   3030
      Begin VB.OptionButton opWhichToClose 
         Caption         =   "Let me choosse"
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   1950
         Value           =   -1  'True
         Width           =   1920
      End
      Begin VB.OptionButton opWhichToClose 
         Caption         =   "Close them all!"
         Height          =   255
         Index           =   1
         Left            =   315
         TabIndex        =   17
         Top             =   2220
         Width           =   2130
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmDetModeWizard.frx":0000
         Top             =   -30
         Width           =   2850
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   180
      ScaleHeight     =   2625
      ScaleWidth      =   3030
      TabIndex        =   11
      Tag             =   "2"
      Top             =   3480
      Width           =   3030
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   1485
         Left            =   15
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmDetModeWizard.frx":008E
         Top             =   750
         Width           =   2880
      End
      Begin VB.TextBox txtPopUp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         HideSelection   =   0   'False
         Left            =   15
         TabIndex        =   13
         Top             =   300
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Here's the PopUp you selected:"
         Height          =   210
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   2625
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   6780
      ScaleHeight     =   2625
      ScaleWidth      =   3030
      TabIndex        =   8
      Tag             =   "4"
      Top             =   3255
      Width           =   3030
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   2040
         Left            =   15
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmDetModeWizard.frx":0127
         Top             =   -15
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   5490
      ScaleHeight     =   2625
      ScaleWidth      =   3030
      TabIndex        =   5
      Tag             =   "1"
      Top             =   225
      Width           =   3030
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmDetModeWizard.frx":01F8
         Top             =   -30
         Width           =   3180
      End
      Begin VB.OptionButton opMode 
         Caption         =   "It will always remain the same"
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   2025
         Value           =   -1  'True
         Width           =   2850
      End
      Begin VB.OptionButton opMode 
         Caption         =   "It may change"
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   2310
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< &Back"
      Height          =   390
      Left            =   2415
      TabIndex        =   4
      Top             =   3015
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   3015
      Width           =   1200
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >>"
      Height          =   390
      Left            =   3735
      TabIndex        =   2
      Top             =   3015
      Width           =   1200
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   1875
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "0"
      Text            =   "frmDetModeWizard.frx":0317
      Top             =   135
      Width           =   3030
   End
   Begin VB.PictureBox picImage 
      Height          =   2625
      Left            =   120
      ScaleHeight     =   2565
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   135
      Width           =   1515
   End
End
Attribute VB_Name = "frmDetModeWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Step As Integer

Private Sub cmdBack_Click()

    If Step >= 0 Then
        If Step = 4 And opMode(0).Value = True Then Step = Step - 2
        Step = Step - 1
    End If
    ZOrderPage

End Sub

Private Sub ZOrderPage()

    Dim ctrl As Control

    For Each ctrl In Controls
        With ctrl
            If Val(.Tag) = Step Then
                .ZOrder 0
            End If
        End With
    Next ctrl
    
    cmdBack.Enabled = Step > 0
    If Step < 4 Then
        cmdNext.Caption = "&Next >>"
    Else
        cmdNext.Caption = "&Finish"
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdNext_Click()

    If Step <= 4 Then
        If Step = 1 And opMode(0).Value = True Then Step = Step + 2
        Step = Step + 1
    End If
    If Step = 5 Then
        NewCustomEntry
    End If
    ZOrderPage

End Sub

Private Sub NewCustomEntry()

    If opMode(0).Value Then
        SelectedPopUps.PopUp.ListPopUps.Add frmMain.lvMonitor.SelectedItem.Text, True, 0, FullText
        frmMain.RefreshSelectedList
        cmdCancel_Click
    Else
        If txtPopUp.SelText = "" Then
            If MsgBox("You must select some of the PopUp's title in order to complete the wizard", vbRetryCancel, "Invalid Wizard Settings") = vbCancel Then
                cmdCancel_Click
            Else
                Step = 2
                ZOrderPage
            End If
        Else
            If opWhichToClose(0).Value Then
                SelectedPopUps.PopUp.ListPopUps.Add txtPopUp.SelText, True, 0, ContentsAsk
            Else
                SelectedPopUps.PopUp.ListPopUps.Add txtPopUp.SelText, True, 0, ContentsClose
            End If
            frmMain.RefreshSelectedList
            cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim ctrl As Control

    Width = 5160
    Height = 3930
    
    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
    txtPopUp.Text = frmMain.lvMonitor.SelectedItem.Text
    
    For Each ctrl In Controls
        With ctrl
            If Val(.Tag) > 0 Then
                .Left = txtInfo.Left
                .Top = txtInfo.Top
            End If
        End With
    Next ctrl
    
    Step = 0
    ZOrderPage

End Sub
