VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSync 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synchronize"
   ClientHeight    =   3180
   ClientLeft      =   5160
   ClientTop       =   4275
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5400
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   2820
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4395
      TabIndex        =   6
      Top             =   1710
      Width           =   870
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4395
      TabIndex        =   5
      Top             =   1200
      Width           =   870
   End
   Begin VB.Frame frmOperations 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4080
      Begin VB.CheckBox chkDownload 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Download available PopUp's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   1155
         Width           =   2640
      End
      Begin VB.CheckBox chkUpload 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Upload my PopUp's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   345
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This will download all the available popups from the server and merge them with yours."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   405
         TabIndex        =   4
         Top             =   1425
         Width           =   3345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This will upload your black list to the server and make it available to others."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   405
         TabIndex        =   3
         Top             =   615
         Width           =   3195
      End
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Idle)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   8
      Top             =   2550
      Width           =   450
   End
   Begin VB.Image imgToggleProgress 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   60
      Picture         =   "frmSync.frx":0000
      Stretch         =   -1  'True
      Top             =   2190
      Width           =   240
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
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
      Left            =   300
      TabIndex        =   7
      Top             =   2190
      Width           =   795
   End
End
Attribute VB_Name = "frmSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim ViewingProgress As Boolean
Dim ErrorConnecting As Boolean
Dim LoggingOff As Boolean
Dim DoneDownload As Boolean
Dim strData As String
Dim Cancel As Boolean

Private Sub chkDownload_Click()

    cmdStart.Enabled = chkDownload.Value Or chkUpload.Value

End Sub

Private Sub chkUpload_Click()

    cmdStart.Enabled = chkDownload.Value Or chkUpload.Value
    If chkUpload.Value Then
        If GetSetting(App.EXEName, "Preferences", "DontAskAgain", 0) = 1 Then
            Ok2Upload = -GetSetting(App.EXEName, "Preferences", "OK2Upload", 0)
        Else
            frmWarning.Show vbModal
        End If
        chkUpload.Value = Abs(Ok2Upload)
    End If

End Sub

Private Sub cmdCancel_Click()

    Cancel = True
    If inetCtrl.StillExecuting Then inetCtrl.Cancel
    Unload Me

End Sub

Private Sub cmdStart_Click()

    frmOperations.Enabled = False
    cmdStart.Enabled = False
    
    cmdCancel.Caption = "Cance&l"

    If chkUpload.Value = vbChecked Then
        UploadPopUps
        If chkDownload.Value = vbChecked Then
            MsgBox "Click the OK button to start the Download operation", vbOKOnly + vbInformation, "PopUp Killer"
        End If
    End If
    
    If chkDownload.Value = vbChecked Then
        DownloadPopUps
    End If
    
    cmdCancel.Caption = "C&lose"
    
    lblAction.Caption = "(idle)"
    
End Sub

Private Sub UploadPopUps()

    On Error Resume Next

    Dim i As Integer
    Dim ff As Integer
    Dim localFile As String
    Dim remoteFile As String
    Dim regUser As String
    Dim volName As String
    Dim drvSerial As String
    
    localFile = Chr(34) + App.Path + "\popups.dat" + Chr(34)
    Kill localFile
    
    regUser = NoSpaces(LCase(QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")))
    rgbGetVolume "c:\", volName, drvSerial
    remoteFile = regUser + ".dat"
    
    With frmMain.lstTitles
        pbProgress.Min = 0
        pbProgress.Max = .ListCount
        
        ff = FreeFile
        Open Mid(localFile, 2, Len(localFile) - 2) For Output As ff
            For i = 1 To .ListCount
                pbProgress.Value = i
                lblAction.Caption = "Preparing PopUp's List (" & Int(i / .ListCount * 100) & ")"
                Print #ff, .List(i - 1)
                DoEvents
            Next i
        Close ff
    End With
    
    ErrorConnecting = False
    With inetCtrl
        .URL = "ftp://xfx.net"
        .UserName = "puk_user"
        .Password = "kqfzzy2536"
        .RequestTimeout = 30
        Err.Clear
        .Execute , "PUT " + localFile + " " + remoteFile
        Do
            DoEvents
        Loop Until Not .StillExecuting _
                    Or ErrorConnecting _
                    Or Err.Number <> 0 _
                    Or Cancel
                    
        If Cancel Then Exit Sub
        
        If ErrorConnecting Then
            MsgBox "Error: " & .ResponseCode & vbCrLf + .ResponseInfo, vbOKCancel + vbCritical, "Error"
            Exit Sub
        ElseIf Err.Number > 0 Then
            MsgBox "Error: " & Err.Number & vbCrLf + Err.Description, vbOKCancel + vbCritical, "Error"
            Exit Sub
        Else
            lblAction.Caption = "Launching Browser..."
        End If
        .Execute , "CLOSE"
    End With
    
    MsgBox "The Upload operation completed succesfuly. PupUp Killer will now open your browser to finish the operation", vbInformation + vbOKOnly, "PopUp Killer"
    Shell "start http://xfx.net/Programs/utilities/popupkiller/confirm.phtml?user=" + regUser, vbHide
    
End Sub

Private Function NoSpaces(str As String) As String

    Dim m As String
    Dim i As Integer
    
    For i = 1 To Len(str)
        m = Mid(str, i, 1)
        If m = " " Then
            m = "_"
        End If
        NoSpaces = NoSpaces + m
    Next i
    
End Function

Private Function uuEncode(str As String) As String

    Dim i As Integer
    Dim m As String
    
    For i = 1 To Len(str)
        m = Mid(str, i, 1)
        If m < "0" Or _
            (m > "9" And m < "A") Or _
            (m > "Z" And m < "a") Or _
            m > "z" Then
            m = "%" + CStr(Hex(Asc(m)))
        End If
        uuEncode = uuEncode + m
    Next i
    
End Function

Private Sub DownloadPopUps()

    Dim i As Integer
    Dim PopUp As String
    Dim Exists As Boolean
    Dim j As Integer

    ErrorConnecting = False
    DoneDownload = False
    strData = ""

    inetGet.Execute
    Do
        DoEvents
    Loop Until ErrorConnecting Or DoneDownload
    
    lblAction.Caption = "Merging PopUps..."
    
    Do While InStr(strData, vbCrLf) And Not Cancel
        PopUp = Left(strData, InStr(strData, vbCrLf) - 1)
        Exists = False
        For i = 0 To frmMain.lstTitles.ListCount - 1
            If frmMain.lstTitles.List(i) = PopUp Then
                Exists = True
                Exit For
            End If
        Next i
        If Len(PopUp) And Not Exists Then
            SelectedPopUps.PopUp.ListPopUps.Add PopUp, True, 0
            j = j + 1
        End If
        strData = Mid(strData, Len(PopUp) + 3)
    Loop
    
    If Cancel Then Exit Sub
    
    MsgBox "A total of " & j & " PopUps have been added to your black list", vbOKOnly + vbInformation, "PopUp Killer"
    frmMain.RefreshSelectedList
    
End Sub

Private Sub Form_Load()

    Cancel = False

    Height = 2895
    
    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
End Sub

Private Sub imgToggleProgress_Click()

    Select Case ViewingProgress
        Case True
            Height = 2895
            imgToggleProgress.Picture = LoadResPicture(103, vbResIcon)
        Case False
            Height = 3585
            imgToggleProgress.Picture = LoadResPicture(104, vbResIcon)
    End Select
    
    ViewingProgress = Not ViewingProgress

End Sub

Private Sub inetCtrl_StateChanged(ByVal State As Integer)

    Select Case State
        Case icResolvingHost
            lblAction.Caption = "Resolving Host..."
        Case icConnecting
            lblAction.Caption = "Connecting..."
        Case icError
            ErrorConnecting = True
        Case icReceivingResponse
            If Not LoggingOff Then
                lblAction.Caption = "Uploading PopUps..."
            End If
        Case icResponseCompleted
            lblAction.Caption = "Operation succesful, Logging Off..."
            LoggingOff = True
    End Select
    
End Sub

Private Sub inetGet_StateChanged(ByVal State As Integer)

    Dim vtData As Variant

    Select Case State
        Case icResolvingHost
            lblAction.Caption = "Resolving Host..."
        Case icConnecting
            lblAction.Caption = "Connecting..."
        Case icError
            ErrorConnecting = True
        Case icResponseCompleted
            vtData = inetGet.GetChunk(1024, icString)
            DoEvents

            Do While Len(vtData)
                lblAction.Caption = "Receiving: " & Int(Len(strData) / 1024) & "KB"
                strData = strData & vtData
                vtData = inetGet.GetChunk(1024, icString)
                DoEvents
            Loop
            
            DoneDownload = True
    
    End Select

End Sub

Private Sub lblProgress_Click()

    imgToggleProgress_Click

End Sub
