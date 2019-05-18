VERSION 5.00
Begin VB.Form frmProxy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proxy Setup"
   ClientHeight    =   1980
   ClientLeft      =   5505
   ClientTop       =   6390
   ClientWidth     =   4410
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1365
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4230
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Top             =   825
         Width           =   735
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   450
         Width           =   1995
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Use a Proxy Server"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   210
         Left            =   855
         TabIndex        =   6
         Top             =   870
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   210
         Left            =   555
         TabIndex        =   4
         Top             =   495
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   3480
      TabIndex        =   1
      Top             =   1530
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2505
      TabIndex        =   0
      Top             =   1530
      Width           =   825
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    SaveSetting App.EXEName, "Proxy", "Use", chkEnable.Value
    SaveSetting App.EXEName, "Proxy", "Address", txtAddress.Text
    SaveSetting App.EXEName, "Proxy", "Port", txtPort.Text
    
    Unload Me

End Sub

Private Sub Form_Load()

    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
    SkinMe Me, True
    SetupCharset Me
    
    chkEnable.Value = GetSetting(App.EXEName, "Proxy", "Use", 0)
    txtAddress.Text = GetSetting(App.EXEName, "Proxy", "Address", "")
    txtPort.Text = GetSetting(App.EXEName, "Proxy", "Port", "")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SkinMe Me, False

End Sub
