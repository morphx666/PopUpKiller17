VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E9AEDB6D-4F13-11D3-9DB5-444553540000}#1.0#0"; "xFXVerCheck.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Version Check"
   ClientHeight    =   3285
   ClientLeft      =   3390
   ClientTop       =   4350
   ClientWidth     =   6810
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
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6810
   Begin VB.PictureBox picHyperlink 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   15
      ScaleHeight     =   375
      ScaleWidth      =   3240
      TabIndex        =   13
      Top             =   2385
      Visible         =   0   'False
      Width           =   3240
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "for information about this version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   16
         Top             =   90
         Width           =   2400
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         Caption         =   "here"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   435
         MouseIcon       =   "frmUpdate.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   75
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Click"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   75
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2790
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   15
      Width           =   480
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "HIDE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2325
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   15
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   270
      Left            =   480
      TabIndex        =   9
      Top             =   1830
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   3525
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   15
      Width           =   3240
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   2250
      TabIndex        =   4
      Top             =   2835
      Width           =   1020
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   390
      Left            =   1080
      TabIndex        =   3
      Top             =   2835
      Width           =   1020
   End
   Begin xFXVerCheck.VerCheck VerCheckCtrl 
      Left            =   2565
      Top             =   585
      _ExtentX        =   820
      _ExtentY        =   767
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Connecting"
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   660
      Width           =   930
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Select destination folder"
      Enabled         =   0   'False
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   2175
      Width           =   1995
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Waiting for user"
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   405
      Width           =   1305
   End
   Begin VB.Image imgPointer 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   210
      Picture         =   "frmUpdate.frx":0594
      Stretch         =   -1  'True
      Top             =   420
      Width           =   210
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Downloading"
      Enabled         =   0   'False
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   6
      Top             =   1590
      Width           =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   135
      X2              =   2900
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2900
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Getting update information"
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   1
      Top             =   1215
      Width           =   2235
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Logging to server"
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   0
      Top             =   930
      Width           =   1440
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim curStep As Integer
Dim RemoteFile As typeSetupFile
Dim DownloadedBytes As Long

Const strAppName = "PopUp Killer"
Const strSAppName = "pukSetup17.zip"
Const strSAppLink = "http://software.xfx.net/utilities/popupkiller/index.html"

Private VerCheckHadError As Boolean
Private WithEvents cBDlg As BrowseDialog
Attribute cBDlg.VB_VarHelpID = -1

Private Sub cmdClose_Click()

    If Not cmdStart.Enabled Then
        If MsgBox("Are you sure you want to cancel the operation?" + vbCrLf + "If you cancel now, you can continue at a later time", vbQuestion + vbYesNo) = vbYes Then
            VerCheckCtrl.CanceOperation
        Else
            Exit Sub
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub MovePointer(idx As Integer)

    With imgPointer
        .Top = lblStatus(idx).Top + lblStatus(idx).Height / 2 - .Height / 2
    End With

End Sub

Private Sub cmdHide_Click()

    Me.SetFocus
    DoEvents

    WindowState = vbMinimized

End Sub

Private Sub cmdMore_Click()

    On Error Resume Next

    Me.SetFocus
    DoEvents

    Select Case Width
        Case Is > 4000
            Width = 3405
        Case Is < 4000
            Width = 6930
    End Select

End Sub

Private Sub cmdStart_Click()

    cmdStart.Enabled = False

    With RemoteFile
        .Path = "/"
        .Name = strSAppName
    End With
    
    curStep = curStep + 1
    MovePointer curStep
    
    With VerCheckCtrl
        .GetSetupInfo RemoteFile
    End With
    
    If Format(RemoteFile.Date, "Short Date") > Format(FileDateTime(App.Path + "\" + App.EXEName + ".exe"), "Short Date") Then
        picHyperlink.Visible = True
        GetUpdate
    Else
        MsgBox "You already have the latest version of " + strAppName, vbInformation, "Version Checker"
    End If
    
    cmdStart.Enabled = True
    
    MovePointer 0

End Sub

Private Sub GetUpdate()

    If MsgBox("A new version of " + strAppName + " is available!" + vbCrLf + "Do you want to download it now?", vbQuestion + vbYesNo, "Version Checker") = vbYes Then
        lblStatus(4).Enabled = True
        lblStatus(5).Enabled = True
        pbDownload.Enabled = True
        pbDownload.Value = 0
        If VerCheckHadError Then
            VerCheckHadError = False
        Else
            VerCheckCtrl.Download RemoteFile
        End If
    Else
        Exit Sub
    End If

End Sub

Private Sub Form_Load()

    cmdMore_Click
    
    DoEvents

    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With

End Sub

Private Sub lblLink_Click()

    Shell "start " + strSAppLink, vbHide

End Sub

Private Sub VerCheckCtrl_DownloadDone(Forced As Boolean)

    If Forced Then
        Unload Me
    Else
        SaveFile
    End If

End Sub

Private Sub SaveFile()

    Dim sFolder As Folder
    Dim sPath As String
    
    MovePointer 5

    Set cBDlg = New BrowseDialog
    With cBDlg
        .AllowNewFolder = True
        .Prompt1 = "Select the folder to store the downloaded file"
        If .Browse Then
            Set sFolder = .SelectedFolder
            FileCopy VerCheckCtrl.GetTempDir + RemoteFile.Name, sFolder.FullPath + "\" + RemoteFile.Name
            MsgBox "The file has been saved at:" + vbCrLf + sFolder.FullPath + "\" + RemoteFile.Name, vbInformation + vbOKOnly
        End If
        Kill VerCheckCtrl.GetTempDir + RemoteFile.Name
        lblStatus(4).Caption = "Downloading"
        
        lblStatus(4).Enabled = False
        lblStatus(5).Enabled = False
        pbDownload.Enabled = False
    End With

End Sub

Private Sub VerCheckCtrl_DownloadStatus(curPacket As Long)

    Dim val As String

    DownloadedBytes = DownloadedBytes + curPacket
    pbDownload.Value = (DownloadedBytes / RemoteFile.Size) * 100
    
    Select Case DownloadedBytes
        Case Is < 1024
            val = CStr(DownloadedBytes) + " Bytes"
        Case 1024 To 1048576 / 2
            val = Format(DownloadedBytes / 1024, "0.00") + " KB"
        Case Is > 1048576 / 2
            val = Format(DownloadedBytes / 1048576, "0.00") + " MB"
    End Select
    
    lblStatus(4).Caption = "Downloading: " + val

End Sub

Private Sub VerCheckCtrl_Error(Description As String)

    MsgBox Description
    VerCheckCtrl.CanceOperation
    
    MovePointer 0
    cmdStart.Enabled = True
    
    VerCheckHadError = True

End Sub

Private Sub VerCheckCtrl_ExtendedStatus(strStatus As String)

    txtLog.Text = txtLog.Text + strStatus + vbCrLf
    txtLog.SelStart = Len(txtLog.Text)

End Sub

Private Sub VerCheckCtrl_Status(strStatus As String)

    Select Case strStatus
        Case "Connecting"
            MovePointer 1
        Case "Logging in"
            MovePointer 2
        Case "Getting Data"
            MovePointer 3
        Case "Downloading"
            MovePointer 4
    End Select

End Sub
