VERSION 5.00
Begin VB.Form frmWarning 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warning"
   ClientHeight    =   2595
   ClientLeft      =   7215
   ClientTop       =   6255
   ClientWidth     =   5415
   Icon            =   "frmWarning.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDontAskAgain 
      Caption         =   "Don't ask me again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   885
      TabIndex        =   6
      Top             =   2250
      Width           =   2190
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3525
      TabIndex        =   5
      Top             =   2130
      Width           =   840
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4485
      TabIndex        =   4
      Top             =   2130
      Width           =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   885
      X2              =   3165
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   885
      X2              =   3165
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to continue?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   885
      TabIndex        =   3
      Top             =   1695
      Width           =   3360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This information will be kept private and no one, besides the webmaster of xfx.net, will have access to this information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   885
      TabIndex        =   2
      Top             =   870
      Width           =   4170
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE USER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   885
      TabIndex        =   1
      Top             =   570
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The upload process will be logged in the server with your Windows registered name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   885
      TabIndex        =   0
      Top             =   135
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmWarning.frx":0442
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()

    Ok2Upload = False
    Unload Me

End Sub

Private Sub cmdYes_Click()

    Ok2Upload = True
    Unload Me

End Sub

Private Sub Form_Load()
    
    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
    SetupCharset Me
    
    chkDontAskAgain.Value = GetSetting(App.EXEName, "Preferences", "DontAskAgain", 0)
    Ok2Upload = -GetSetting(App.EXEName, "Preferences", "OK2Upload", 0)
    lblUserName.Caption = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
    If Command = "/xfx" Then lblUserName.Caption = "xfxjs"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveSetting App.EXEName, "Preferences", "OK2Upload", Abs(Ok2Upload)
    SaveSetting App.EXEName, "Preferences", "DontAskAgain", chkDontAskAgain.Value

End Sub
