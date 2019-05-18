VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "HREF.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{07ED89F4-A824-4EC6-AE85-E2F8FFD0CCF1}#1.0#0"; "xFXSysTray.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PopUp Killer "
   ClientHeight    =   7875
   ClientLeft      =   4080
   ClientTop       =   3150
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   10110
   Begin VB.PictureBox picSync 
      Height          =   3975
      Left            =   6210
      ScaleHeight     =   3915
      ScaleWidth      =   3765
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   3825
      Begin VB.CommandButton cmdProxy 
         Caption         =   "&Proxy Setup"
         Height          =   360
         Left            =   60
         TabIndex        =   18
         Top             =   3390
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Height          =   360
         Left            =   2895
         TabIndex        =   12
         Top             =   3405
         Width           =   810
      End
      Begin VB.CheckBox chkUpload 
         Caption         =   "Upload my PopUp's"
         Height          =   285
         Left            =   105
         TabIndex        =   9
         Top             =   105
         Width           =   1950
      End
      Begin VB.CheckBox chkDownload 
         Caption         =   "Download available PopUp's"
         Height          =   285
         Left            =   105
         TabIndex        =   8
         Top             =   915
         Width           =   2640
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
         Height          =   360
         Left            =   1995
         TabIndex        =   6
         Top             =   3405
         Width           =   810
      End
      Begin VB.Line Line4 
         X1              =   75
         X2              =   3660
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   90
         X2              =   3675
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label lblSkipped 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skipped:"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   525
         TabIndex        =   28
         Top             =   2730
         Width           =   705
      End
      Begin VB.Label lblDownloaded 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded:"
         Height          =   210
         Left            =   150
         TabIndex        =   27
         Top             =   2325
         Width           =   1080
      End
      Begin VB.Label lblAdded 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Added:"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   630
         TabIndex        =   26
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label lblAction 
         AutoSize        =   -1  'True
         Caption         =   "(Idle)"
         Height          =   210
         Left            =   1350
         TabIndex        =   14
         Top             =   1935
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   210
         Left            =   645
         TabIndex        =   13
         Top             =   1935
         Width           =   585
      End
      Begin VB.Line LineW 
         BorderColor     =   &H80000005&
         X1              =   165
         X2              =   3750
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line lineB 
         X1              =   150
         X2              =   3735
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label lblUpload 
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
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   3765
      End
      Begin VB.Label lblDownload 
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
         ForeColor       =   &H00808080&
         Height          =   480
         Left            =   120
         TabIndex        =   10
         Top             =   1185
         Width           =   3765
      End
   End
   Begin VB.PictureBox picAbout 
      Height          =   4605
      Left            =   2475
      ScaleHeight     =   4545
      ScaleWidth      =   4995
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdGetSN 
         Caption         =   "Get Serial Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3585
         TabIndex        =   41
         Top             =   4065
         Width           =   945
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   135
         TabIndex        =   37
         Top             =   2865
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   53
      End
      Begin VB.TextBox txtRegKey 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4275
         Width           =   2250
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2565
         TabIndex        =   17
         Top             =   4065
         Width           =   945
      End
      Begin href.uchref1 hrefSupport 
         Height          =   315
         Left            =   135
         TabIndex        =   32
         Top             =   2430
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         Caption         =   "Click here to jump to the online Help"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16711680
         URL             =   "http://software.xfx.net/utilities/popupkiller/tsguide/index.html"
      End
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   135
         TabIndex        =   38
         Top             =   3420
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   135
         TabIndex        =   39
         Top             =   3975
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   53
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Official beta-tester and online blacklist administrator: Bruce R. H."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   43
         Top             =   2955
         Width           =   4110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered to"
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
         Left            =   180
         TabIndex        =   36
         Top             =   4065
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine v2.0 code optimization has been made possible thanks to Jeroen van der Ham"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   3510
         Width           =   3360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configuration Guide:"
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
         Left            =   180
         TabIndex        =   34
         Top             =   2235
         Width           =   1485
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.0.0"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4275
         TabIndex        =   16
         Top             =   1710
         Width           =   435
      End
      Begin VB.Image imgLogo 
         Height          =   2145
         Left            =   0
         Picture         =   "frmMain.frx":0442
         Top             =   0
         Width           =   4800
      End
      Begin VB.Image imgExtender 
         Height          =   915
         Left            =   4485
         Picture         =   "frmMain.frx":3C14
         Stretch         =   -1  'True
         Top             =   2895
         Width           =   315
      End
   End
   Begin VB.PictureBox picPopUps 
      Height          =   4395
      Left            =   4185
      ScaleHeight     =   4335
      ScaleWidth      =   4470
      TabIndex        =   1
      Top             =   1575
      Width           =   4530
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "frmMain.frx":66CC
         Top             =   2955
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvTitles 
         Height          =   1500
         Left            =   75
         TabIndex        =   22
         Top             =   330
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   2646
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chTitle"
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvMonitor 
         Height          =   1545
         Left            =   75
         TabIndex        =   2
         ToolTipText     =   "To add a popup to the black list, just double click it"
         Top             =   2640
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   2725
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chTitle"
            Text            =   "PopUp's Titles"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblListDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PopUps that will be closed (black list)"
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
         Left            =   75
         TabIndex        =   4
         Top             =   75
         Width           =   3435
      End
      Begin VB.Label lblMonitor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detected PopUps"
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
         Left            =   75
         TabIndex        =   3
         Top             =   2400
         Width           =   1590
      End
   End
   Begin VB.Timer tmrMonReg 
      Interval        =   1500
      Left            =   1815
      Top             =   5730
   End
   Begin VB.PictureBox picLog 
      Height          =   2460
      Left            =   120
      ScaleHeight     =   2400
      ScaleWidth      =   3420
      TabIndex        =   19
      Top             =   3060
      Width           =   3480
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2940
         TabIndex        =   42
         Top             =   45
         Width           =   555
      End
      Begin MSComctlLib.ListView lvLog 
         Height          =   990
         Left            =   60
         TabIndex        =   20
         Top             =   210
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1746
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chTitle"
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "chEngine"
            Text            =   "Engine"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "chCount"
            Text            =   "Count"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "chBLEntry"
            Text            =   "Black List Entry"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PopUps that have been closed"
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
         Left            =   75
         TabIndex        =   21
         Top             =   75
         Width           =   2790
      End
   End
   Begin VB.PictureBox picExclude 
      Height          =   2835
      Left            =   5790
      ScaleHeight     =   2775
      ScaleWidth      =   3435
      TabIndex        =   29
      Top             =   4380
      Width           =   3495
      Begin MSComctlLib.ListView lvExclude 
         Height          =   1545
         Left            =   0
         TabIndex        =   30
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   2725
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chTitle"
            Text            =   "PopUp's Titles"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PopUps that will not be closed"
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
         Left            =   75
         TabIndex        =   31
         Top             =   75
         Width           =   2820
      End
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   2760
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin xFXSysTray.xFXcSysTray trayCtrl 
      Left            =   1830
      Top             =   6675
      _ExtentX        =   450
      _ExtentY        =   450
      InTray          =   0   'False
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.PictureBox picBanned 
      Height          =   1095
      Left            =   3075
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   6525
      Width           =   1215
      Begin VB.CommandButton cmdUploadBP 
         Caption         =   "&Upload"
         Height          =   330
         Left            =   105
         TabIndex        =   25
         Top             =   660
         Width           =   900
      End
      Begin MSComctlLib.ListView lvBanned 
         Height          =   300
         Left            =   45
         TabIndex        =   24
         Top             =   45
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chTitle"
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilTabs 
      Left            =   2415
      Top             =   6465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":683D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C91
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7951
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A219
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrFlashDelay 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1275
      Top             =   5790
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   7590
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2037
            MinWidth        =   2028
            Key             =   "pSelPopUps"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2037
            MinWidth        =   2028
            Key             =   "pState"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
            Key             =   "pProgress"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsPanels 
      Height          =   4905
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   8652
      TabWidthStyle   =   1
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   882
      ImageList       =   "ilTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PopUps"
            Key             =   "tsPopUps"
            Object.ToolTipText     =   "PopUps"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Exclusions"
            Key             =   "tsExclusions"
            Object.ToolTipText     =   "Exclusions"
            ImageVarType    =   2
            ImageIndex      =   8
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Banned"
            Key             =   "tsBanned"
            Object.ToolTipText     =   "Banned"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Synchronize"
            Key             =   "tsSync"
            Object.ToolTipText     =   "Synchronize"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            Key             =   "tsLog"
            Object.ToolTipText     =   "Log"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "tsHelp"
            Object.ToolTipText     =   "Help"
            ImageVarType    =   2
            ImageIndex      =   7
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "tsAbout"
            Object.ToolTipText     =   "About"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrMonitor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   750
      Top             =   5775
   End
   Begin VB.Timer tmrGetActivePopUps 
      Interval        =   500
      Left            =   165
      Top             =   5850
   End
   Begin InetCtlsObjects.Inet inetCtrl 
      Left            =   60
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      RemoteHost      =   "software.xfx.net"
      URL             =   "http://xfx.net/utilities/popupkiller/getpopups.php3"
      Document        =   "/utilities/popupkiller/getpopups.php3"
   End
   Begin InetCtlsObjects.Inet inetGet 
      Left            =   750
      Top             =   6585
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      RemoteHost      =   "software.xfx.net"
      URL             =   "http://software.xfx.net/utilities/popupkiller/getpopups.php3"
      Document        =   "/utilities/popupkiller/getpopups.php3"
   End
   Begin VB.Menu mnuItm 
      Caption         =   "PopUp"
      Begin VB.Menu mnuItmAdd 
         Caption         =   "&Add..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuItmSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItmBanIt 
         Caption         =   "&Ban it!"
      End
      Begin VB.Menu mnuItmExclude 
         Caption         =   "&Exclude"
      End
      Begin VB.Menu mnuItmSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItmEdit 
         Caption         =   "&Edit..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuItmRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuItmSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItmFind 
         Caption         =   "&Find..."
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Options"
      Begin VB.Menu mnuOpOpen 
         Caption         =   "&Open PopUp Killer"
      End
      Begin VB.Menu mnuOpSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpPreferences 
         Caption         =   "&Preferences..."
      End
      Begin VB.Menu mnuOpSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpDisable 
         Caption         =   "&Disable"
      End
      Begin VB.Menu mnuOpSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpBL 
         Caption         =   "Black List"
         Begin VB.Menu mnuOpBLLoad 
            Caption         =   "Load..."
         End
         Begin VB.Menu mnuOpBLMerge 
            Caption         =   "Merge..."
         End
         Begin VB.Menu mnuOpBLSync 
            Caption         =   "Synchronize..."
         End
         Begin VB.Menu mnuOpBLSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpBLSaveAs 
            Caption         =   "Save As..."
         End
         Begin VB.Menu mnuOpBLSep02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpBLWipe 
            Caption         =   "Wipe"
         End
      End
      Begin VB.Menu mnuOpSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuOpCheckUpdate 
         Caption         =   "&Check for Update"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuOpSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpMin 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuOpExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mntBL 
      Caption         =   "Black List"
      Begin VB.Menu mntBLLoad 
         Caption         =   "&Load..."
      End
      Begin VB.Menu mntBLMerge 
         Caption         =   "&Merge..."
      End
      Begin VB.Menu mnuOpSync 
         Caption         =   "&Synchronize..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mntBLSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mntBLSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mntBLSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mntBLWipe 
         Caption         =   "&Wipe"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Main Code
Dim OverrideOnline As Boolean
Dim MenuVisible As Boolean
Dim Modifyed As Boolean
Dim Refreshing As Boolean
Dim DisableBanned As Boolean
Dim CancelSizeSave As Boolean
Dim MonitorChanged As Boolean
Dim Exiting As Boolean
Dim HaltMonitor As Boolean
Dim Flashes As Integer
Dim sWidth As Integer
Dim sHeight As Integer

'Sync Code
Dim ViewingProgress As Boolean
Dim ErrorConnecting As Boolean
Dim LoggingOff As Boolean
Dim strData As String
Dim Cancel As Boolean
Dim IsSync As Boolean
Dim GettingBanned As Boolean
Dim StartMerging As Boolean

Dim INIFile As String
Dim BannedFile As String

Dim MonitorBusy As Boolean
Dim EnumeratorBusy As Boolean

Dim ForceFocusOnMenu As Boolean

Dim ProgramState As Integer

Private Function CheckWin2K() As Boolean

    Dim IpInts() As String
    Dim qRes As String
    Dim i As Integer
    
    IpInts = Split(QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters\Adapters\NdisWanIp", "IpConfig"), Chr(0))
    
    For i = LBound(IpInts) To UBound(IpInts)
        qRes = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" + IpInts(i), "DhcpIPAddress")
        If qRes <> "" And qRes <> "0.0.0.0" Then
            CheckWin2K = True
            Exit Function
        End If
    Next i

End Function

Private Function CheckWin9x() As Boolean

    CheckWin9x = (QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\RemoteAccess", "Remote Connection") <> 0)

End Function

Private Sub GetActivePopUps()

    Dim i As Long
    Dim ii As Integer
    Dim Title As String
    Dim cName As String
    Dim mItem As ListItem
    Dim p As Integer
    Dim lch() As Long
    
    GetWindowHandles
    
    For Each mItem In lvMonitor.ListItems
        mItem.Tag = 0
    Next mItem
    
    For i = 1 To UBound(lHandles)
        cName = clsName(lHandles(i))
        If Left(LCase(cName), 13) = "afx:400000:b:" Or Left(LCase(cName), 13) = "afx:400000:30" Then
            If InStr(wTitle(lHandles(i)), "Netscape") Then cName = "FNETCRAP4"
        End If
        Select Case cName
            Case "IEFrame", "FNETCRAP4", "MozillaWindowClass", "CabinetWClass" ', "ThunderRT6FormDC"
                Title = wTitle(lHandles(i))
                p = InStrRev(Title, " - ")
                If p Or cName = "FNETCRAP4" Then
                    If p Then
                        Title = Left(Title, p - 1)
                        AddTitle Title, lHandles(i)
                    End If
                    
                    Title = GetEditText(lHandles(i), "Edit")
                    If Title <> "" Then AddTitle Title, lHandles(i)
                End If
            Case "BLDOPERA"
                GetChildWindowHandles lHandles(i)
                lch = lcHandles
                For ii = 1 To UBound(lch)
                    If clsName(lch(ii)) = "BLD_ObjWin" Then
                        Title = wTitle(lch(ii))
                        If Title <> "..." Then
                            AddTitle Title, lch(ii)
                        
                            Title = GetEditText(lch(ii), "Edit")
                            If Title <> "" Then AddTitle Title, lch(ii)
                        End If
                    End If
                Next ii
            Case "OUIWINDOW"
                Title = wTitle(lHandles(i))
                AddTitle Title, lHandles(i)
                Title = GetEditText(lHandles(i), "ComboBox")
                If Title <> "" Then AddTitle Title, lHandles(i)
            Case "Opera Main Window"
                GetChildWindowHandles lHandles(i)
                lch = lcHandles
                For ii = 1 To UBound(lch)
                    If clsName(lch(ii)) = "BrowserWindowMDI" Then
                        Title = wTitle(lch(ii))
                        Title = Mid(Title, 2)
                        If Title <> ".." Then
                            AddTitle Title, lch(ii)
                    
                            Title = GetEditText(lch(ii), "ComboBox")
                            If Title <> "" Then AddTitle Title, lch(ii)
                        End If
                    End If
                Next ii
            Case "MSN6 Window"
                AddTitle wTitle(lHandles(i)), lHandles(i)
            Case "MozillaUIWindowClass"
                Title = wTitle(lHandles(i))
                p = InStrRev(Title, " - ")
                If p Then
                    Title = Left(Title, p - 1)
                    AddTitle Title, lHandles(i)
                End If
            Case "#32770"
                Title = wTitle(lHandles(i))
                Select Case Title
                    Case "[JavaScript Application]", "Home Page", "Microsoft Internet Explorer", "Explorer User Prompt" ', "Confirm", "Alert"
                        AddTitle "Dialog: " + GetDlgText(lHandles(i)), lHandles(i)
                    Case Else
                        If Right(Title, 12) = " - NeoPlanet" Then
                            AddTitle Mid(Title, 1, InStrRev(Title, " - ") - 1), lHandles(i)
                            
                            Title = GetEditText(lHandles(i), "ComboBox")
                            If Title <> "" Then AddTitle Title, lHandles(i)
                        Else
                            If Right(Title, 12) = " - NeoPlanet" Then
                                
                            Else
                                If FindIEServer(lHandles(i)) Then
                                    AddTitle "Dialog: " + Title, lHandles(i)
                                End If
                            End If
                        End If
                        'AddTitle "Dialog: " + GetDlgText(lHandles(i)), lHandles(i)
                End Select
            Case "Internet Explorer_TridentDlgFrame"
                Title = wTitle(lHandles(i))
                If Title <> "Find" Then
                    AddTitle "Dialog: " + Title, lHandles(i)
                End If
        End Select
    Next i
    
ReStart:
    For Each mItem In lvMonitor.ListItems
        If mItem.Tag = 0 Then
            lvMonitor.ListItems.Remove mItem.Index
            MonitorChanged = True
            'GoTo ReStart
            Exit For
        End If
    Next mItem
    txtInfo.Visible = (lvMonitor.ListItems.Count = 0)
    
End Sub

Private Function FindIEServer(lh As Long) As Boolean

    Dim i As Integer

    GetChildWindowHandles lh
    For i = 1 To UBound(lcHandles)
        If clsName(lcHandles(i)) = "Internet Explorer_Server" Then
            FindIEServer = True
        End If
    Next i

End Function

Private Function GetEditText(lh As Long, ctrl As String) As String

    Dim i As Integer
    Dim Title As String

    GetChildWindowHandles lh
    For i = 1 To UBound(lcHandles)
        If clsName(lcHandles(i)) = ctrl Then
            Title = Space(SendMessageLong(lcHandles(i), WM_GETTEXTLENGTH, 0, ByVal 0))
            If Title <> "" Then
                SendMessageLong lcHandles(i), WM_GETTEXT, Len(Title) + 1, ByVal Title
                If InStr(Title, "://") Then
                    GetEditText = Title
                    Exit For
                End If
            End If
        End If
    Next i

End Function

Friend Sub AddTitle(Title As String, TitleHandle As Long)

    On Error Resume Next

    Dim FixedTitle As String
    Dim dummy As Integer
    Dim nItem As ListItem
    
    With lvMonitor.ListItems
        dummy = .Item("k" + Title).Index
        If dummy = 0 Then
            Set nItem = .Add(, "k" + Title, Title)
            MonitorChanged = True
        Else
            Set nItem = .Item(dummy)
        End If
    End With
    nItem.Tag = TitleHandle

End Sub

Private Function ParseTitle(Title As String) As String

    If InStr(Title, " - Microsoft Internet Explorer") <> 0 _
                Or InStr(Title, " - Mozilla") <> 0 _
                Or InStr(Title, "NeoPlanet") <> 0 _
                Or InStr(Title, " - Netscape") <> 0 _
                Or InStr(Title, "[JavaScript Application]") <> 0 _
                Or (Preferences.UseCustomIETitle And (InStr(Title, Preferences.CustomIETitle) <> 0)) Then
                If InStr(Title, " - ") Then
                    ParseTitle = Left$(Title, InStrRev(Title, " - ") - 1)
                End If
    End If

End Function

Private Function NewParser(str As String) As String

    Dim i As Integer
    
    i = InStrRev(str, "-")
    If i > 0 Then
        NewParser = Left$(str, i - 2)
    Else
        NewParser = str
    End If
    
End Function

Friend Sub AddSelected(Title As String, Enabled As Boolean, Optional SkipExistsTest As Boolean)

    On Error GoTo ShowError

    Dim nItem As ListItem
    Dim OkToAdd As Boolean
    
    If Not SkipExistsTest Then
        OkToAdd = (Not IsInList(lvBanned, Title, True)) And (Not IsInList(lvExclude, Title, True)) And (Not IsInList(lvTitles, Title, True))
    Else
        OkToAdd = True
    End If
    If OkToAdd Then
        Set nItem = lvTitles.ListItems.Add(, , Title)
        With nItem
            .Checked = Enabled
        End With
    End If
    
    Exit Sub
    
ShowError:
    MsgBox "An error has occurred trying to add a PopUp to the black list" + vbCrLf + Err.Description, vbOKOnly
    
End Sub

Private Sub GetBannedPopUps()

    On Error Resume Next

    Dim ff As Integer
    Dim str As String
    
    lvBanned.ListItems.Clear
    ff = FreeFile
    If Dir(BannedFile) <> "" Then
        Open BannedFile For Input As #ff
            Do Until EOF(ff)
                Line Input #ff, str
                lvBanned.ListItems.Add , , str
            Loop
        Close #ff
    End If

End Sub

Private Sub SaveBannedPopUps()

    Dim ff As Integer
    Dim bItem As ListItem
    
    If Dir(BannedFile) <> "" Then Kill BannedFile
    
    ff = FreeFile
    Open BannedFile For Output As #ff
        For Each bItem In lvBanned.ListItems
            Print #ff, bItem.Text
        Next bItem
    Close #ff

End Sub

Private Sub LoadPopUps()

    On Error Resume Next

    Dim SavedTitles As String
    Dim sepPos As Integer
    
    Dim curTitle As String
    Dim curState As Boolean
    Dim i As Long
    Dim ff As Integer
    
    Dim curLine As String
    
    lvTitles.ListItems.Clear
    
    INIFile = App.Path + "\popups.ini"
    BannedFile = App.Path + "\banned.ini"
    GetBannedPopUps
    
    ff = FreeFile
    
    trayCtrl.TrayTip = "PopUp Killer is: Loading Black List..."
    
    If Dir(INIFile) <> "" Then
        Open INIFile For Input As ff
            Do Until EOF(ff)
                Line Input #ff, curLine
                curTitle = Left$(curLine, Len(curLine) - 2)
                curState = -Val(Right$(curLine, 1))
                AddSelected curTitle, curState, True
            Loop
        Close ff
    End If
    
    UpdateStatusBar
    
End Sub

Private Sub LoadExclusions()

    On Error Resume Next

    Dim SavedTitles As String
    Dim sepPos As Integer
    
    Dim curTitle As String
    Dim curState As Boolean
    Dim i As Long
    Dim ff As Integer
    
    Dim curLine As String
    
    lvExclude.ListItems.Clear
    
    ff = FreeFile
    
    If Dir(App.Path + "\expopups.ini") <> "" Then
        Open App.Path + "\expopups.ini" For Input As ff
            Do Until EOF(ff)
                Line Input #ff, curLine
                curTitle = Left$(curLine, Len(curLine) - 2)
                curState = -Val(Right$(curLine, 1))
                lvExclude.ListItems.Add , , curTitle
            Loop
        Close ff
    End If
    
End Sub

Private Sub UpdateStatusBar()

    With sbInfo
        .Panels("pSelPopUps").Text = lvTitles.ListItems.Count & " PopUps"
    End With
    
    AutoColumnsSize lvTitles

End Sub

Private Sub chkDownload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors
    lblDownload.ForeColor = vbBlack

End Sub

Private Sub chkUpload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors
    lblUpload.ForeColor = vbBlack

End Sub

Private Sub cmdCancel_Click()

    Cancel = True

End Sub

Private Sub cmdClear_Click()

    lvLog.ListItems.Clear
    lvLog.SetFocus

End Sub

Private Sub cmdGetSN_Click()

    Clipboard.Clear
    Clipboard.SetText EncryptHDSerial, vbCFText
    
    MsgBox "The Serial Number has been copied to your clipboard", vbInformation + vbOKOnly, "PopUp Killer Serial Number"

End Sub

Private Sub cmdProxy_Click()

    frmProxy.Show vbModal

End Sub

Private Sub cmdRegister_Click()

    MsgBox "PopUp Killer is FREE." + vbCrLf + _
            "There's no need for you to register it... but..." + vbCrLf + _
            "if you want to contribute I would really appreciate your donation." + vbCrLf + _
            "The donation is US$7.00 from which we'll get US$5.60." + vbCrLf + _
            "So if you decide to give your donation, please go to" + vbCrLf + _
            "    http://software.xfx.net/utilities/popupkiller/register.htm" + vbCrLf + _
            "and follow the instructions", vbOKOnly + vbInformation

    With frmSerialize
        .lblMsg.Caption = "Enter your name and the license code you recevied by email. Note that the license code is case sensitive."
        .Show vbModal
    End With
    
    If IsSerialized2(intAppname, "PopUpKiller") Then
        MsgBox "Thank you very much for registering this product!" + vbCrLf + "Your support will help us keep improving current titles and working on new ones that for sure will be of your liking.", vbExclamation + vbOKOnly
        txtRegKey.Text = GetSetting(App.EXEName, "Config", "User")
        cmdRegister.Visible = Not IsSerialized2(intAppname, "PopUpKiller")
        cmdGetSN.Visible = Not IsSerialized2(intAppname, "PopUpKiller")
    Else
        MsgBox "The license code you entered is invalid. Make sure you have typed it exactly as appears in the email your received. Also note that the product's license code is case sensitive." + vbCrLf + "Click on the Register button again to re-enter the code.", vbExclamation + vbOKOnly
    End If

End Sub

Private Sub cmdUploadBP_Click()

    On Error Resume Next

    Dim i As Integer
    Dim ff As Integer
    
    ErrorConnecting = False
    With inetCtrl
        .URL = "ftp://software.xfx.net"
        .UserName = "puk_user"
        .Password = "the_password"
        .RequestTimeout = 15
        If GetSetting(App.EXEName, "Proxy", "Use", 0) = 1 Then
            .AccessType = icNamedProxy
            .Proxy = GetSetting(App.EXEName, "Proxy", "Address") + ":" + GetSetting(App.EXEName, "Proxy", "Port")
        Else
            .AccessType = icUseDefault
        End If
        Err.Clear
        .Execute , "PUT """ + BannedFile + """ banned.ini"
        Do
            DoEvents
        Loop Until Not .StillExecuting _
                    Or ErrorConnecting _
                    Or Err.Number <> 0 _
                    Or Cancel
                    
        If Cancel Then Exit Sub
        
        MsgBox "Done.."
        
        If ErrorConnecting Then
            MsgBox "Error: " & .ResponseCode & vbCrLf + .ResponseInfo, vbOKCancel + vbCritical, "Error"
            Exit Sub
        ElseIf Err.Number > 0 Then
            MsgBox "Error: " & Err.Number & vbCrLf + Err.Description, vbOKCancel + vbCritical, "Error"
            Exit Sub
        End If
        .Execute , "CLOSE"
    End With
    
End Sub

Private Sub ResetColors()

    lblDownload.ForeColor = &H808080
    lblUpload.ForeColor = &H808080

End Sub

Private Sub Form_Initialize()

    If App.PrevInstance And Command <> "/forceclose" Then
        Beep
        SaveSetting App.EXEName, "Preferences", "ForceShow", 1
        End
    Else
        If Command = "/forceclose" Then
            SaveSetting App.EXEName, "Preferences", "ForceShow", 2
            End
        Else
            SaveSetting App.EXEName, "Preferences", "ForceShow", 0
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    If Exiting Then End
    
    If InStr(Command, "/debug") = 0 Then rhSubClass Me.hwnd, 344, 404
    
    CancelSizeSave = True
    
    HelpFile = App.Path + "\puk.chm"
    
    SaveSetting App.EXEName, "Config", "InstallPath", App.Path
    
    Left = GetSetting(App.EXEName, "Preferences", "Left", Screen.Width / 3)
    Top = GetSetting(App.EXEName, "Preferences", "Top", Screen.Height / 3)
    Width = GetSetting(App.EXEName, "Preferences", "Width", 5085)
    Height = GetSetting(App.EXEName, "Preferences", "Height", 5580)
    
    CancelSizeSave = False
    
    With Preferences
        .AutoStart = GetSetting(App.EXEName, "Preferences", "AutoStart", True)
        .JumpToHomePage = GetSetting(App.EXEName, "Preferences", "JumpToHomePage", False)
        .HomePage = GetSetting(App.EXEName, "Preferences", "HomePage", QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page"))
        
        .FlashIcon = GetSetting(App.EXEName, "Preferences", "Flash", False)
        .ShowIcon = GetSetting(App.EXEName, "Preferences", "ShowInTray", 1)
        .ShowTabsText = GetSetting(App.EXEName, "Preferences", "ShowTabsText", True)
        .ScrambleText = GetSetting(App.EXEName, "Preferences", "ScrambleText", False)
        .SoundFX = GetSetting(App.EXEName, "Preferences", "Sound", 1)
        
        .Smart = GetSetting(App.EXEName, "Preferences", "Smart", 0)
        .SafeMode = GetSetting(App.EXEName, "Preferences", "SafeSmart", 0)
        .SmartSensitivity = GetSetting(App.EXEName, "Preferences", "SmartSens", 120000)
        .Limit = GetSetting(App.EXEName, "Preferences", "LimitMode", False)
        .LimitNum = GetSetting(App.EXEName, "Preferences", "LimitWindows", 3)
        .FGeo = GetSetting(App.EXEName, "Preferences", "FGeo", False)
        
        .UseCustomIETitle = GetSetting(App.EXEName, "Preferences", "UseCustomIETitle", False)
        .CustomIETitle = GetSetting(App.EXEName, "Preferences", "CustomIETitle", QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title"))
        .UseCustomIETitle = .UseCustomIETitle And (.CustomIETitle <> "")
        
        .DisableShortcutKeys = GetSetting(App.EXEName, "Preferences", "DisableSCK", False)
        .DisableWildcards = GetSetting(App.EXEName, "Preferences", "DisableWC", False)
        
        .Changed = True
    End With
    
    With trayCtrl
        Set .TrayIcon = LoadResPicture(101, vbResIcon)
        .TrayTip = "PopUp Killer is: Idle"
    End With
    
    If QueryValue(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\PUK\") <> "PopUp Killer" Then
        CreateNewKey "AppEvents\Schemes\Apps\PUK\", HKEY_CURRENT_USER
        SetKeyValue HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\PUK\", , "PopUp Killer"
        CreateNewKey "AppEvents\Schemes\Apps\PUK\PopUp Anihilation", HKEY_CURRENT_USER
        CreateNewKey "AppEvents\Schemes\Apps\PUK\PopUp Anihilation\.Current", HKEY_CURRENT_USER
        SetKeyValue HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\PUK\PopUp Anihilation\.Current", , App.Path + "\SOUNDS\THUNDER.WAV"
    End If
    
    sWidth = Screen.Width / Screen.TwipsPerPixelX
    sHeight = Screen.Height / Screen.TwipsPerPixelY
    
    mnuOpMin.Enabled = False
    Me.Hide
    trayCtrl.InTray = Preferences.ShowIcon
    
    cmdRegister.Visible = Not IsSerialized2(intAppname, "PopUpKiller")
    cmdGetSN.Visible = Not IsSerialized2(intAppname, "PopUpKiller")
    txtRegKey.Text = GetSetting(App.EXEName, "Config", "User")
    
    picPopUps.ZOrder 0
    
    LoadPopUps
    LoadExclusions
    RefFGeoLib
    
    DisableCloseButton Me
    
    On Error Resume Next
    FontCharSet = Printer.Font.Charset
    If Err.Number <> 0 Then
        FontCharSet = Me.Font.Charset
    End If
    Err.Clear
    On Error GoTo 0
    
    SkinMe Me, True
    SetupCharset Me
    
    UpdatePreferences
    
    If InStr(Command, "/xfx") = 0 Then
        DisableBanned = True
        mnuItmBanIt.Visible = False
        tsPanels.Tabs.Remove tsPanels.Tabs("tsBanned").Index
    End If
    
    MonitorChanged = True
    
End Sub

Private Sub RefFGeoLib()

    On Error Resume Next
    #If DEBUGGEOPLUGIN = 0 Then
    Set FGeoLib = CreateObject("FGeo.clsFGeo")
    #End If

End Sub

Private Sub SavePopUps(FileName As String)

    Dim Title As String
    Dim ff As Integer
    Dim pItem As ListItem
    Dim SourceControl As ListView
    
    If Not Modifyed And FileName = INIFile Then Exit Sub
    
    If Dir(FileName) <> "" Then Kill FileName
    ff = FreeFile
    
    trayCtrl.TrayTip = "Please Wait... Saving PopUps!"
    
    If FileName = INIFile Then
        Set SourceControl = lvTitles
    Else
        Set SourceControl = lvExclude
    End If
    
    Open FileName For Output As ff
        For Each pItem In SourceControl.ListItems
            Title = pItem.Text
            Title = Title + Chr(255) & Abs(pItem.Checked)
            Print #ff, Title
        Next pItem
    Close ff
    
    Modifyed = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode >= 2 Then
        Exiting = True
        SkinMe Me, False
        SaveProgramSettings
    End If

End Sub

Private Sub Form_Resize()

    Dim ctrl As Control
    Dim cTop As Long
    
    If Me.WindowState = vbMinimized Then
        mnuOpMin_Click
        Exit Sub
    End If
    
    cTop = GetClientTop(Me.hwnd)
    
    With tsPanels
        .Top = 25
        .Width = Width - 125
        .Height = Height - sbInfo.Height - cTop - 25
    End With
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is PictureBox Then
            With ctrl
                .BorderStyle = 0
                .Top = tsPanels.ClientTop + 30
                .Left = tsPanels.ClientLeft
                .Width = tsPanels.ClientWidth
                .Height = tsPanels.ClientHeight - 90
            End With
        ElseIf TypeOf ctrl Is CommandButton Then
            ctrl.Top = tsPanels.ClientHeight - tsPanels.ClientTop
        ElseIf TypeOf ctrl Is Line Then
            ctrl.X1 = tsPanels.ClientLeft
            ctrl.X2 = tsPanels.ClientWidth - 195
        End If
    Next ctrl
    
    cmdCancel.Move tsPanels.ClientWidth - cmdCancel.Width - 275, tsPanels.Height - 890
    cmdStart.Move cmdCancel.Left - cmdStart.Width - 275, cmdCancel.Top
    cmdProxy.Top = cmdCancel.Top
    cmdRegister.Move 2600, 4070
    cmdGetSN.Move 3600, 4070
    cmdClear.Move lvLog.Left + lvLog.Width - cmdClear.Width, lvLog.Top - cmdClear.Height
    
    cmdUploadBP.Move cmdCancel.Left, cmdCancel.Top
    
    lblMonitor.Top = tsPanels.ClientHeight - (400 + lvMonitor.Height)
    With lvMonitor
        .Width = tsPanels.ClientWidth - 225
        .ColumnHeaders(1).Width = .Width - 375
        .Top = lblMonitor.Top + 250
        .Height = 2500
        txtInfo.Move .Left + 30, .Top + 30, .Width - 60, .Height - 60
    End With
    
    lvExclude.Move 60, lvTitles.Top, picExclude.Width - 155, picExclude.Height - 300
    lvLog.Move 60, lvTitles.Top, picExclude.Width - 155, picExclude.Height - 300
    lvBanned.Move 60, lvTitles.Top, picExclude.Width - 155, cmdUploadBP.Top - 300
    
    AutoColumnsSize lvLog
    AutoColumnsSize lvExclude
    lvTitles.ColumnHeaders(1).Width = lvTitles.Width - 375
    lvBanned.ColumnHeaders(1).Width = lvBanned.Width - 375
    
    With lvTitles
        .Width = tsPanels.ClientWidth - 255
        .Height = lblMonitor.Top - 415
    End With
    
    If Not CancelSizeSave Then
        SaveSetting App.EXEName, "Preferences", "Left", Left
        SaveSetting App.EXEName, "Preferences", "Top", Top
        SaveSetting App.EXEName, "Preferences", "Width", Width
        SaveSetting App.EXEName, "Preferences", "Height", Height
    End If
    
    imgLogo.Move tsPanels.Width / 2 - imgLogo.Width / 2, 0
    imgExtender.Move 0, 0, tsPanels.Width, imgLogo.Height
    lblVersion.Left = Width - lblVersion.Width - 400
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    rhUnSubClass Me.hwnd

    If defWindowProc2 Then
        Call SetWindowLong(frmMain.hwnd, GWL_WNDPROC, defWindowProc2)
        defWindowProc2 = 0
    End If

End Sub

Private Sub lvExclude_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim NextItm As ListItem
    
    If KeyCode = 46 Then
        With lvExclude
            If .ListItems.Count Then
                If .SelectedItem.Index <> .ListItems.Count Then
                    Set NextItm = .ListItems(.SelectedItem.Index + 1)
                End If
                mnuItmRemove_Click
                If Not NextItm Is Nothing Then
                    NextItm.Selected = True
                    NextItm.EnsureVisible
                End If
            End If
        End With
    End If

End Sub

Private Sub lvExclude_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not lvExclude.HitTest(X, Y) Is Nothing Then
        lvExclude.HitTest(X, Y).Selected = True
    End If

    If Button = vbRightButton Then
        mnuItmAdd.Enabled = True
        mnuItmEdit.Enabled = True
        mnuItmBanIt.Enabled = False
        mnuItmExclude.Enabled = False
        mnuItmFind.Enabled = True
        mnuItmRemove.Enabled = True
        PopupMenu mnuItm, vbRightButton
    End If
    
End Sub

Private Sub lvLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvLog
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
    
End Sub

Private Sub lvLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not lvLog.HitTest(X, Y) Is Nothing Then
        lvLog.HitTest(X, Y).Selected = True
    End If

    If Button = vbRightButton Then
        mnuItmAdd.Enabled = True
        mnuItmBanIt.Enabled = True
        mnuItmExclude.Enabled = True
        mnuItmFind.Enabled = False
        mnuItmRemove.Enabled = False
        mnuItmEdit.Enabled = False
        PopupMenu mnuItm, vbRightButton
    End If

End Sub

Private Sub lvMonitor_DblClick()

    On Error Resume Next
    
    AddSelected lvMonitor.SelectedItem.Text, True
    MonitorChanged = True
    Modifyed = True

End Sub

Private Sub lvMonitor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        mnuItmAdd.Enabled = True
        mnuItmBanIt.Enabled = False
        mnuItmExclude.Enabled = False
        mnuItmFind.Enabled = False
        mnuItmRemove.Enabled = False
        mnuItmEdit.Enabled = False
        PopupMenu mnuItm, vbRightButton
    End If

End Sub

Private Sub lvTitles_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Modifyed = True And Not Refreshing
    
End Sub

Private Sub lvTitles_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim NextItm As ListItem
    
    If KeyCode = 46 Then
        If lvTitles.ListItems.Count Then
            If lvTitles.SelectedItem.Index <> lvTitles.ListItems.Count Then
                Set NextItm = lvTitles.ListItems(lvTitles.SelectedItem.Index + 1)
            End If
            mnuItmRemove_Click
            If Not NextItm Is Nothing Then
                NextItm.Selected = True
                NextItm.EnsureVisible
            End If
        End If
    End If

End Sub

Private Sub lvTitles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lvTitles.MultiSelect = False
    DoEvents

    If lvTitles.ListItems.Count > 0 And Not (lvTitles.HitTest(X, Y) Is Nothing) Then
        With lvTitles.HitTest(X, Y)
            .EnsureVisible
            .Selected = True
        End With
    End If
    
    lvTitles.MultiSelect = True

    If Button = vbRightButton Then
        mnuItmAdd.Enabled = False
        mnuItmBanIt.Enabled = True
        mnuItmExclude.Enabled = True
        mnuItmFind.Enabled = True
        mnuItmEdit.Enabled = True
        mnuItmRemove.Enabled = True
        PopupMenu mnuItm, vbRightButton
    End If

End Sub

Private Sub lvTitles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    AddSelected Data.GetData(vbCFText), True
    
    Modifyed = True

End Sub

Private Sub mntBLLoad_Click()

    On Error GoTo CancelSub
    
    Dim SrcPath As String
    
    If MsgBox("Loading a Black List will overwrite your current list of popups." + vbCrLf + "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbYes Then
        With cDlg
            .CancelError = True
            .DialogTitle = "Load Black List"
            .Filter = "Black List|*.ini"
            .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNShareAware
            .ShowOpen
            SrcPath = Left$(.FileName, InStrRev(.FileName, "\"))
            FileCopy SrcPath + "popups.ini", App.Path + "\popups.ini"
            FileCopy SrcPath + "expopups.ini", App.Path + "\expopups.ini"
            LoadPopUps
        End With
        Modifyed = False
    End If
    
    MonitorChanged = True
    
CancelSub:

End Sub

Private Sub mntBLMerge_Click()

    On Error GoTo CancelSub
    
    Dim ff As Integer
    Dim curLine As String
    Dim curTitle As String
    Dim curState As Boolean
    
    With cDlg
        .CancelError = True
        .DialogTitle = "Merge Black List"
        .Filter = "Black List|*.ini"
        .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNShareAware
        .ShowOpen
        ff = FreeFile
        Open .FileName For Input As ff
            Do Until EOF(ff)
                Line Input #ff, curLine
                curTitle = Left$(curLine, Len(curLine) - 2)
                If Not (IsInList(lvTitles, curTitle, True) Or IsInList(lvExclude, curTitle, True) Or IsInList(lvBanned, curTitle, True)) Then
                    curState = -Val(Right$(curLine, 1))
                    AddSelected curTitle, curState, False
                    UpdateStatusBar
                    sbInfo.Refresh
                End If
            Loop
        Close ff
    End With
    UpdateStatusBar
    Modifyed = True
    SavePopUps INIFile
    
    MonitorChanged = True
    
CancelSub:

End Sub

Private Sub mntBLSaveAs_Click()

    On Error GoTo CancelSub
    
    Dim DestPath As String

    With cDlg
        .CancelError = True
        .DialogTitle = "Save Black List As"
        .Filter = "Black List"
        .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly
        .FileName = "popups.ini"
        .ShowSave
        DestPath = Left$(.FileName, InStrRev(.FileName, "\"))
        FileCopy App.Path + "\popups.ini", DestPath + "popups.ini"
        FileCopy App.Path + "\expopups.ini", DestPath + "expopups.ini"
    End With
    
CancelSub:

End Sub

Private Sub mntBLWipe_Click()

    If MsgBox("Wiping a Black List will delete your current list of popups." + vbCrLf + "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbYes Then
        lvTitles.ListItems.Clear
        On Error Resume Next
        Kill App.Path + "\popups.ini"
        Modifyed = False
    End If

End Sub

Private Sub mnuHelp_Click()

    RunShellExecute "Open", "hh.exe", HelpFile + "::intro.html", 0, 1

End Sub

Private Sub mnuItmAdd_Click()

    On Error Resume Next

    tmrMonitor.Enabled = False
    
    With frmAddPopUp
        Select Case tsPanels.SelectedItem.Key
            Case "tsPopUps"
                .txtLabel.Text = lvMonitor.SelectedItem.Text
            Case "tsLog"
                .txtLabel.Text = lvLog.SelectedItem.Text
        End Select
        .opList(0).Value = (tsPanels.SelectedItem.Key = "tsPopUps") Or (tsPanels.SelectedItem.Key = "tsLog")
        .opList(1).Value = Not .opList(0).Value
        .Show vbModal
        SavePopUps INIFile
        SavePopUps App.Path + "/expopups.ini"
    End With
    
    MonitorChanged = True
    Modifyed = True
    
    tmrMonitor.Enabled = True

End Sub

Private Sub mnuItmBanIt_Click()

    Select Case tsPanels.SelectedItem.Key
        Case "tsPopUps"
            If lvTitles.ListItems.Count > 0 Then
                lvBanned.ListItems.Add , , lvTitles.SelectedItem.Text
                lvTitles.ListItems.Remove lvTitles.SelectedItem.Index
            End If
        Case "tsLog"
            If lvLog.ListItems.Count > 0 Then
                If IsInList(lvBanned, lvLog.SelectedItem.Text, True) Then
                    MsgBox "The Selected PopUp cannot be added because is a banned popup", vbInformation + vbOKCancel, "Unable to add popup"
                Else
                    lvBanned.ListItems.Add , , lvLog.SelectedItem.Text
                    lvLog.ListItems.Remove lvLog.SelectedItem.Index
                End If
            End If
    End Select
    
    SaveBannedPopUps

End Sub

Private Sub mnuItmEdit_Click()

    Dim HasChanged As Boolean
    Dim NewText As String
    
    tmrMonitor.Enabled = False
    
    Select Case ActiveControl.Name
        Case "lvTitles"
            If lvTitles.SelectedItem Is Nothing Then Exit Sub
            NewText = InputBox("Enter the new name for the selected popup", "Edit PopUp", lvTitles.SelectedItem.Text)
            If NewText <> "" Then
                lvTitles.SelectedItem.Text = NewText
                HasChanged = True
            End If
        Case "lvExclude"
            If lvExclude.SelectedItem Is Nothing Then Exit Sub
            NewText = InputBox("Enter the new name for the selected popup", "Edit PopUp", lvExclude.SelectedItem.Text)
            If NewText <> "" Then
                lvExclude.SelectedItem.Text = NewText
                HasChanged = True
            End If
    End Select
    
    SavePopUps INIFile
    SavePopUps App.Path + "/expopups.ini"
    
    tmrMonitor.Enabled = True
    
    MonitorChanged = True
    Modifyed = True

End Sub

Private Sub mnuItmExclude_Click()

    On Error Resume Next

    Select Case tsPanels.SelectedItem.Key
        Case "tsPopUps"
            If Not IsInList(lvExclude, lvTitles.SelectedItem.Text, True) Then
                lvExclude.ListItems.Add , , lvTitles.SelectedItem.Text
                mnuItmRemove_Click
            End If
        Case "tsLog"
            If Not IsInList(lvExclude, lvLog.SelectedItem.Text, True) Then
                lvExclude.ListItems.Add , , lvLog.SelectedItem.Text
                lvLog.ListItems.Remove lvLog.SelectedItem.Index
            End If
    End Select
    
    SavePopUps App.Path + "\expopups.ini"

End Sub

Private Sub mnuItmFind_Click()

    FindPopUp False

End Sub

Private Sub FindPopUp(SkipPrompt As Boolean)

    Static str As String
    Dim i As Integer
    Dim SecondTry As Boolean
    Dim SrcControl As ListView
    
    Select Case tsPanels.SelectedItem.Key
        Case "tsPopUps"
            Set SrcControl = lvTitles
        Case "tsExclusions"
            Set SrcControl = lvExclude
    End Select
    
    If SrcControl Is Nothing Then Exit Sub
    
    If SrcControl.ListItems.Count = 0 Then Exit Sub
    
    If Not SkipPrompt Then
        str = LCase$(InputBox("Enter a text string to look for", "Find PopUp Title", str))
        If str = "" Then Exit Sub
    End If
    
    str = LCase$(str)
    SrcControl.MultiSelect = False
    With SrcControl.ListItems
        i = SrcControl.SelectedItem.Index + 1
        If i > .Count Then i = 1
Retry:
        For i = i To .Count
            If InStr(LCase$(.Item(i).Text), str) > 0 Then
                .Item(i).EnsureVisible
                .Item(i).Selected = True
                Exit Sub
            End If
        Next i
        If Not SecondTry Then
            i = 1
            SecondTry = True
            GoTo Retry
        End If
    End With
    SrcControl.MultiSelect = True
    
    SrcControl.SetFocus

End Sub

Private Sub mnuItmRemove_Click()

    Dim sItem As ListItem
    Dim SrcControl As ListView
    
    Select Case tsPanels.SelectedItem.Key
        Case "tsPopUps"
            Set SrcControl = lvTitles
        Case "tsExclusions"
            Set SrcControl = lvExclude
    End Select
    
    If SrcControl.ListItems.Count = 0 Then Exit Sub
    
ReStart:
    For Each sItem In SrcControl.ListItems
        If sItem.Selected Then
            SrcControl.ListItems.Remove sItem.Index
            GoTo ReStart
        End If
    Next sItem

    UpdateStatusBar
    
    Modifyed = True

End Sub

Private Sub mnuOpAbout_Click()

    MakeSound
    mnuOpOpen_Click
    tsPanels.Tabs("tsAbout").Selected = True
    
End Sub

Private Sub mnuOpBLLoad_Click()

    mntBLLoad_Click

End Sub

Private Sub mnuOpBLMerge_Click()

    mntBLMerge_Click

End Sub

Private Sub mnuOpBLSaveAs_Click()

    mntBLSaveAs_Click

End Sub

Private Sub mnuOpBLSync_Click()

    mnuOpSync_Click

End Sub

Private Sub mnuOpBLWipe_Click()

    mntBLWipe_Click

End Sub

Private Sub mnuOpCheckUpdate_Click()

    Dim LastState As Boolean

    LastState = trayCtrl.InTray
    trayCtrl.InTray = False
    trayCtrl.InTray = LastState

End Sub

Private Sub mnuOpDisable_Click()

    mnuOpDisable.Checked = Not mnuOpDisable.Checked
    If mnuOpDisable.Checked Then
        With trayCtrl
            Set .TrayIcon = LoadResPicture(105, vbResIcon)
            .TrayTip = "PopUp Killer is: Disabled"
        End With
        sbInfo.Panels("pState").Text = "Disabled"
    Else
        MonitorChanged = True
    End If

End Sub

Private Sub mnuOpExit_Click()

    tmrFlashDelay.Enabled = False
    tmrGetActivePopUps.Enabled = False
    tmrMonitor.Enabled = False
    DoEvents
    SaveProgramSettings
    DoEvents

    Unload Me

End Sub

Private Sub SaveProgramSettings()

    SavePopUps INIFile
    SavePopUps App.Path + "\expopups.ini"
    SaveBannedPopUps
    
    With Preferences
        SaveSetting App.EXEName, "Preferences", "AutoStart", .AutoStart
        SaveSetting App.EXEName, "Preferences", "JumpToHomePage", .JumpToHomePage
        SaveSetting App.EXEName, "Preferences", "HomePage", .HomePage
        
        SaveSetting App.EXEName, "Preferences", "Flash", .FlashIcon
        SaveSetting App.EXEName, "Preferences", "ShowInTray", .ShowIcon
        SaveSetting App.EXEName, "Preferences", "ShowTabsText", .ShowTabsText
        SaveSetting App.EXEName, "Preferences", "ScrambleText", .ScrambleText
        SaveSetting App.EXEName, "Preferences", "Sound", .SoundFX
        
        SaveSetting App.EXEName, "Preferences", "Smart", .Smart
        SaveSetting App.EXEName, "Preferences", "SafeSmart", .SafeMode
        SaveSetting App.EXEName, "Preferences", "SmartSens", .SmartSensitivity
        SaveSetting App.EXEName, "Preferences", "LimitMode", .Limit
        SaveSetting App.EXEName, "Preferences", "LimitWindows", .LimitNum
        SaveSetting App.EXEName, "Preferences", "FGeo", .FGeo
        
        SaveSetting App.EXEName, "Preferences", "UseCustomIETitle", .UseCustomIETitle
        SaveSetting App.EXEName, "Preferences", "CustomIETitle", .CustomIETitle
        
        SaveSetting App.EXEName, "Preferences", "DisableSCK", .DisableShortcutKeys
        SaveSetting App.EXEName, "Preferences", "DisableWC", .DisableWildcards
    End With
    
    If Preferences.AutoStart Then
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "PopUpKiller", App.Path + "\" + App.EXEName + ".EXE"
    Else
        DeleteKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "PopUpKiller"
    End If
    
    trayCtrl.InTray = False
    
End Sub

Private Sub mnuOpMin_Click()

    Me.Hide
    mnuOpMin.Enabled = False
    mnuOpOpen.Enabled = True
    mnuOpBL.Visible = True
    mnuOpSep04.Visible = True

End Sub

Public Sub mnuOpOpen_Click()

    On Error GoTo ExitSub

    If mnuOpOpen.Enabled Then
        mnuOpBL.Visible = False
        mnuOpSep04.Visible = False
        mnuOpMin.Enabled = True
        mnuOpOpen.Enabled = False
        tmrMonitor.Enabled = False
        Me.WindowState = vbNormal
        Me.Show
        If IsSync Then
            tsPanels.Tabs("tsSync").Selected = True
        Else
            tsPanels.Tabs("tsPopUps").Selected = True
        End If
    End If
    
ExitSub:
    
End Sub

Private Sub mnuOpPreferences_Click()

    frmOptions.Show vbModal
    UpdatePreferences
    
    MonitorChanged = True

End Sub

Private Sub UpdatePreferences()

    Dim i As Integer

    With Preferences
        If .Changed Then
            trayCtrl.InTray = .ShowIcon
            lvTitles.Font.Name = IIf(.ScrambleText, "Marlett", Me.Font.Name)
            
            For i = 1 To tsPanels.Tabs.Count
                tsPanels.Tabs(i).Caption = IIf(.ShowTabsText, tsPanels.Tabs(i).ToolTipText, "")
            Next i
            tsPanels.Refresh
            Form_Resize
        End If
    End With

End Sub

Private Sub mnuOpSync_Click()

    SavePopUps INIFile
    mnuOpOpen_Click
    tsPanels.Tabs("tsSync").Selected = True
    
    MonitorChanged = True

End Sub

Private Sub picSync_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetColors

End Sub

Private Sub tmrFlashDelay_Timer()

    Flashes = Flashes + 1
    If Int(Flashes / 2) = Flashes / 2 Then
        Set trayCtrl.TrayIcon = LoadResPicture(106, vbResIcon)
    Else
        Set trayCtrl.TrayIcon = LoadResPicture(102, vbResIcon)
    End If
    
    If Flashes >= 9 Then tmrFlashDelay.Enabled = False

End Sub

Private Sub tmrGetActivePopUps_Timer()

    Static fDis As Integer
    
    If Not Preferences.DisableShortcutKeys Then
        HaltMonitor = (GetAsyncKeyState(&H11) And &H8001) <> 0
        If HaltMonitor And (GetAsyncKeyState(&H10) And &H8001) Then
            If mnuOpDisable.Checked Then
                mnuOpDisable_Click
            Else
                fDis = fDis + 1
                If fDis = 15 Then mnuOpDisable_Click
            End If
        Else
            fDis = 0
        End If
    End If

    Select Case mnuOpDisable.Checked
        Case True
            If ProgramState <> 3 Then
                tmrGetActivePopUps.Interval = 3000
                tmrMonitor.Enabled = False
                Set trayCtrl.TrayIcon = LoadResPicture(101, vbResIcon)
                trayCtrl.TrayTip = "PopUp Killer is: Idle"
                sbInfo.Panels("pState").Text = "Idle"
                txtInfo.Visible = True
                ProgramState = 3
            End If
            Exit Sub
        Case False
            tmrGetActivePopUps.Interval = 100
            tmrMonitor.Enabled = True
            Select Case HaltMonitor
                Case False
                    If ProgramState <> 2 Then
                        Set trayCtrl.TrayIcon = LoadResPicture(102, vbResIcon)
                        trayCtrl.TrayTip = "PopUp Killer is: Active"
                        sbInfo.Panels("pState").Text = "Active"
                        ProgramState = 2
                    End If
                Case True
                    If ProgramState <> 1 Then
                        Set trayCtrl.TrayIcon = LoadResPicture(107, vbResIcon)
                        trayCtrl.TrayTip = "PopUp Killer is: Halted"
                        sbInfo.Panels("pState").Text = "Halted"
                        ProgramState = 1
                    End If
            End Select
    End Select

    EnumeratorBusy = True
    
    If ForceFocusOnMenu Then
        SetMenuFocus Me
        ForceFocusOnMenu = False
    End If

    If Not mnuOpDisable.Checked Then
        GetActivePopUps
        If Preferences.FGeo Then Try2FGeo
    End If
    
    EnumeratorBusy = False
    
End Sub

Private Sub tmrMonitor_Timer()

    Dim i As Integer
    Dim tItem As ListItem
    Dim mItem As ListItem
    Dim blItem As ListItem
    
    If mnuOpDisable.Checked Then Exit Sub
    If MonitorBusy Or HaltMonitor Then Exit Sub
    If Not MonitorChanged Then Exit Sub
    
    MonitorBusy = True
    MonitorChanged = False
    
    On Error GoTo chkErr
    
    If Preferences.Limit Then
        For i = lvMonitor.ListItems.Count To (Preferences.LimitNum + 1) * 2 Step -1
            KillThePopUp lvMonitor.ListItems(i), "Limit Mode"
        Next i
    End If
    
    MatchPopups

chkErr:

    DoEvents
    
    MonitorBusy = False
    
End Sub

Private Sub MatchPopups()

    Dim mItem As ListItem
    Dim HasKilled As Boolean

    For Each mItem In lvMonitor.ListItems
        If Not IsInList(lvBanned, mItem.Text, Not Preferences.DisableWildcards) Then
            If Not IsInList(lvExclude, mItem.Text, Not Preferences.DisableWildcards) Then
                If IsInList(lvTitles, mItem.Text, Not Preferences.DisableWildcards) Then
                    KillThePopUp mItem, "Black List"
                    HasKilled = True
                Else
                    If SmartPopUpDetection(Val(mItem.Tag), mItem.Text) Then
                        KillThePopUp mItem, "Smart!"
                        HasKilled = True
                    End If
                End If
            End If
        End If
    Next mItem
    
    tmrMonitor.Interval = IIf(HasKilled, 10, 500)

End Sub

Private Sub Try2FGeo()

    Dim mItm As ListItem

    If Not FGeoLib Is Nothing Then
        For Each mItm In lvMonitor.ListItems
            If InStr(1, mItm.Text, "geocities", vbTextCompare) Or _
               InStr(1, mItm.Text, "yahoo", vbTextCompare) Then
                If FGeoLib.Scan Then
                    Add2Log Nothing, "GeoCities BOX Plugin"
                End If
            End If
        Next mItm
    End If

End Sub

Private Sub SelectKilledPopUp(pTitle As String)

    On Error Resume Next

    With lvTitles
        .SetFocus
        .MultiSelect = False
        With .FindItem(pTitle, lvwText, , lvwWhole)
            .Selected = True
            .EnsureVisible
        End With
        .MultiSelect = True
    End With
    
End Sub

Private Sub KillThePopUp(p As ListItem, Engine As String)

    Dim keys() As String
    Static ki As Integer
    keys = Split("{ESC}|n|c", "|")
    
    'PostMessage p.Tag, WM_DESTROY, 0&, 0&
    If clsName(p.Tag) = "#32770" Then
        BringWindowToTop p.Tag
        SetActiveWindow p.Tag
        SetFocusAPI p.Tag
        DoEvents
        SendKeys keys(ki)
        ki = ki + 1: If ki > UBound(keys) Then ki = 0
    End If
    PostMessage p.Tag, WM_CLOSE, 0&, 0&
    
    FlashIcon
    MakeSound
    Add2Log p, Engine
    
    If lvMonitor.ListItems.Count = 1 And Preferences.JumpToHomePage Then
        RunShellExecute "Open", Preferences.HomePage, "", 0, 1
    End If
    
    MonitorChanged = True
    
End Sub

Private Sub Add2Log(p As ListItem, Engine As String)

    Dim nItem As ListItem
    Dim Title As String
    
    On Error Resume Next
    
    Title = p.Text

    With lvLog
        Set nItem = .FindItem(Title)
        If nItem Is Nothing Then
            Set nItem = lvLog.ListItems.Add(, , Title)
        End If
    End With
    
    With nItem
        .SubItems(1) = Engine
        .SubItems(2) = Val(.SubItems(2)) + 1
        If Engine = "Black List" Then .SubItems(3) = lvTitles.SelectedItem.Text
        .SubItems(4) = Time
    End With
    
    AutoColumnsSize lvLog

End Sub

Private Sub FlashIcon()

    If Preferences.FlashIcon Then
        Flashes = 0
        tmrFlashDelay.Enabled = True
    End If

End Sub

Private Sub MakeSound()

    If Preferences.SoundFX Then
        PlaySound QueryValue(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\PUK\PopUp Anihilation\.Current"), _
                    0&, SND_ASYNC Or SND_FILENAME Or SND_NOWAIT
    End If

End Sub

Private Function SmartPopUpDetection(hwnd As Long, Title As String) As Boolean

    On Error Resume Next

    Dim wRect As RECT
    Dim h As Long
    Dim w As Long
    Dim a As Long
    Dim cName As String
    
    Static IgnoredTitles() As String
    Dim i As Integer
    
    tmrMonitor.Enabled = False
    
    If Preferences.Smart Then
        cName = clsName(hwnd)
        If cName = "#32770" Then GoTo SkipOthers
        If cName = "Internet Explorer_TridentDlgFrame" Then GoTo SkipOthers
        GetWindowRect hwnd, wRect
        With wRect
            h = (.Bottom - .Top)
            w = (.Right - .Left)
            a = h * w
            If cName = "BrowserWindowMDI" Or cName = "BLD_ObjWin" Then
                a = a * 3
            End If
            If a < Preferences.SmartSensitivity And a > 5000 Then
                SmartPopUpDetection = True
            End If
        End With
    Else
        SmartPopUpDetection = False
    End If
    
    'List of Safe Sites that use PopUps
    If SmartPopUpDetection Then
    
        If Not lvMonitor.FindItem("Microsoft Windows Update - providing critical updates, security fixes, and software downloads", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Microsoft Windows Update", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("http://www.blink.com", lvwText, , lvwPartial) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Blink", lvwText, , lvwPartial) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Barclays Internet Banking - Welcome", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("MULTIMEDIA-STUDIO URECH c/o Chreon", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("AtomFilms: Instant Entertainment 04", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Welcome to Netscape WebMail", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("i-drive.com", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Dialpad.com", lvwText, , lvwWhole) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If Not lvMonitor.FindItem("Net@ddress", lvwText, , lvwPartial) Is Nothing Then
            SmartPopUpDetection = False
            GoTo SkipOthers
        End If
        
        If SmartPopUpDetection Then
            
            Modifyed = True
        
            If SmartPopUpDetection And Preferences.SafeMode Then
            
                i = UBound(IgnoredTitles)
                If Err.Number = 0 Then
                    For i = 0 To UBound(IgnoredTitles)
                        If IgnoredTitles(i) = Title Then
                            SmartPopUpDetection = False
                            Exit For
                        End If
                    Next i
                End If
            
                If SmartPopUpDetection Then
                    DoEvents
                    With frmSafeSmart
                        SafeSmartAns = Invalid
                        SafeSmartHwnd = hwnd
                        .lblMsg.Caption = "The Smart! engine has detected that """ + Title + """ is a possible PopUp window." + vbCrLf + "Are you sure you want to close it?"
                        .Show
                        Do While SafeSmartAns = Invalid
                            DoEvents
                        Loop
                    End With
                    
                    SmartPopUpDetection = (SafeSmartAns = CloseIt) Or _
                                          (SafeSmartAns = CloseIt_and_Add2BlackList)
                                          
                    If SafeSmartAns = IgnoreIt Or SafeSmartAns = IgnoreIt_and_Add2IgnoreList Then
                        ReDim Preserve IgnoredTitles(UBound(IgnoredTitles) + 1)
                        If Err.Number Then ReDim IgnoredTitles(0)
                        IgnoredTitles(UBound(IgnoredTitles)) = Title
                    End If
                                          
                    If SafeSmartAns = CloseIt_and_Add2BlackList Then
                        AddSelected Title, True, False
                    End If
                    If SafeSmartAns = IgnoreIt_and_Add2IgnoreList Then
                        lvExclude.ListItems.Add , , Title
                    End If
                End If
            End If
        End If
        
SkipOthers:
        
    End If
    
    tmrMonitor.Enabled = True
    
End Function

Private Sub tmrMonReg_Timer()

    If Not Preferences.ShowIcon Then
        If GetSetting(App.EXEName, "Preferences", "ForceShow", 0) = 1 Then
            SaveSetting App.EXEName, "Preferences", "ForceShow", 0
            mnuOpOpen_Click
        End If
        
        If GetSetting(App.EXEName, "Preferences", "ForceShow", 0) = 2 Then
            SaveSetting App.EXEName, "Preferences", "ForceShow", 0
            tmrFlashDelay.Enabled = False
            tmrGetActivePopUps.Enabled = False
            tmrMonitor.Enabled = False
            DoEvents
            End
        End If
    End If

End Sub

Private Sub trayCtrl_MouseDblClick(Button As Integer, Id As Long)

    If Button = vbLeftButton Then
        mnuOpOpen_Click
    End If
    
    On Error Resume Next
    Me.SetFocus

End Sub

Private Sub trayCtrl_MouseUp(Button As Integer, Id As Long)

    If Button = vbRightButton And Not Me.Visible Then
        ForceFocusOnMenu = True
        PopupMenu mnuOp, vbLeftButton, , , mnuOpOpen
    End If

End Sub

Private Sub tsPanels_Click()

    Dim pic As PictureBox
    
    If tsPanels.SelectedItem.Key = "tsHelp" Then
        mnuHelp_Click
        Exit Sub
    End If

    picPopUps.Visible = False
    picSync.Visible = False
    picLog.Visible = False
    picAbout.Visible = False
    picBanned.Visible = False
    
    mnuItmAdd.Enabled = False
    mnuItmBanIt.Enabled = False
    mnuItmEdit.Enabled = False
    mnuItmExclude.Enabled = False
    mnuItmFind.Enabled = False
    mnuItmRemove.Enabled = False
    
    Select Case tsPanels.SelectedItem.Key
        Case "tsPopUps"
            Set pic = picPopUps
            mnuItmAdd.Enabled = True
            mnuItmBanIt.Enabled = True
            mnuItmEdit.Enabled = True
            mnuItmExclude.Enabled = True
            mnuItmFind.Enabled = True
            mnuItmRemove.Enabled = True
        Case "tsSync"
            Set pic = picSync
        Case "tsAbout"
            Set pic = picAbout
            lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
        Case "tsLog"
            Set pic = picLog
            mnuItmAdd.Enabled = True
            mnuItmBanIt.Enabled = True
            mnuItmExclude.Enabled = True
        Case "tsBanned"
            Set pic = picBanned
        Case "tsExclusions"
            Set pic = picExclude
    End Select
    
    pic.Visible = True
    pic.ZOrder 0
    
    On Error Resume Next
    
    Me.SetFocus

End Sub

Private Sub chkDownload_Click()

    cmdStart.Enabled = chkDownload.Value Or chkUpload.Value

End Sub

Private Sub chkUpload_Click()

    If chkUpload.Value = vbChecked And InStr(Command, "/xfx") = 0 Then
        MsgBox "Uploading your black list is not possible right now due to problems with our server." + vbCrLf + "You may select to download the popups listed on our server though", vbInformation + vbOKOnly, "Operation not permitted"
        chkUpload.Value = vbUnchecked
        Exit Sub
    End If

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

Private Sub cmdStart_Click()

    Dim LastState As Boolean
    
    IsSync = True

    LastState = mnuOpDisable.Checked
    mnuOpDisable.Checked = True
    DoEvents
    lvTitles.Visible = False
    tsPanels.Enabled = False
    cmdStart.Enabled = False
    cmdCancel.Enabled = True

    If chkUpload.Value = vbChecked Then
        UploadPopUps
        If chkDownload.Value = vbChecked Then
            MsgBox "Click the OK button to start the Download operation", vbOKOnly + vbInformation, "PopUp Killer"
        End If
    End If
    
    If chkDownload.Value = vbChecked Then
        DownloadBanned
        DownloadPopUps
        UpdateStatusBar
        lblAction.Caption = "Saving PopUps..."
        SavePopUps INIFile
    End If
    
    cmdCancel.Enabled = False
    lblAction.Caption = "(idle)"
    
    mnuOpDisable.Checked = LastState
    lvTitles.Visible = True
    tsPanels.Enabled = True
    
    IsSync = False
    
End Sub

Private Sub UploadPopUps()

    On Error Resume Next

    Dim i As Integer
    Dim ff As Integer
    Dim localFile As String
    Dim RemoteFile As String
    Dim regUser As String
    Dim volName As String
    Dim drvSerial As String
    Dim pukVer As String
    
    localFile = Chr(34) + App.Path + "\popups.dat" + Chr(34)
    Kill localFile
    
    regUser = NoSpaces(LCase$(QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")))
    If InStr(Command, "/xfx") <> 0 Then regUser = "xfxjs"
    rgbGetVolume "c:\", volName, drvSerial
    RemoteFile = regUser + ".dat"
    
    With lvTitles.ListItems
                
        ff = FreeFile
        Open Mid$(localFile, 2, Len(localFile) - 2) For Output As ff
            For i = 1 To .Count
                lblAction.Caption = "Preparing PopUp's List (" & Int(i / .Count * 100) & ")"
                Print #ff, .Item(i).Text
                DoEvents
            Next i
        Close ff
    End With
    
    Cancel = False
    ErrorConnecting = False
    With inetCtrl
        .URL = "ftp://software.xfx.net"
        .UserName = "puk_user"
        .Password = "the_password"
        .RequestTimeout = 15
        If GetSetting(App.EXEName, "Proxy", "Use", 0) = 1 Then
            .AccessType = icNamedProxy
            .Proxy = GetSetting(App.EXEName, "Proxy", "Address") + ":" + GetSetting(App.EXEName, "Proxy", "Port")
        Else
            .AccessType = icUseDefault
        End If
        Err.Clear
        .Execute , "PUT " + localFile + " " + RemoteFile
        Do
            DoEvents
        Loop Until Not .StillExecuting _
                    Or ErrorConnecting _
                    Or Err.Number <> 0 _
                    Or Cancel
                    
        If Cancel Then Exit Sub
        
        If ErrorConnecting Then
            MsgBox "Error: " & .ResponseCode & vbCrLf + .ResponseInfo, vbOKOnly + vbCritical, "Error"
            Exit Sub
        ElseIf Err.Number > 0 Then
            MsgBox "Error: " & Err.Number & vbCrLf + Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            lblAction.Caption = "Launching Browser..."
        End If
        .Execute , "CLOSE"
    End With
    
    MsgBox "The Upload operation completed succesfuly. PupUp Killer will now open your browser to finish the operation", vbInformation + vbOKOnly, "PopUp Killer"
    pukVer = App.Major & Format(App.Minor, "00") & Format(App.Revision, "000")
    RunShellExecute "Open", "http://software.xfx.net/utilities/popupkiller/confirm.php?pukver=" & pukVer & "&user=" + regUser, "", 0, 1
    
End Sub

Private Function NoSpaces(str As String) As String

    Dim m As String
    Dim i As Integer
    
    For i = 1 To Len(str)
        m = Mid$(str, i, 1)
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
        m = Mid$(str, i, 1)
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
    Dim j As Integer
    Dim e As Integer
    Dim Total As Long
    Dim t As Integer
    Dim vtData As Variant
    Dim DoneDownload As Boolean
    
    lblAdded.Caption = "Added:"
    lblDownloaded.Caption = "Downloaded:"
    lblSkipped.Caption = "Skipped:"
    
    Modifyed = True
    Cancel = False
    ErrorConnecting = False
    DoneDownload = False
    strData = ""
    
    StartMerging = False
    With inetGet
        If GetSetting(App.EXEName, "Proxy", "Use", 0) = 1 Then
            .AccessType = icNamedProxy
            .Proxy = GetSetting(App.EXEName, "Proxy", "Address") + ":" + GetSetting(App.EXEName, "Proxy", "Port")
        Else
            .AccessType = icUseDefault
        End If
        .Protocol = icHTTP
        .RemotePort = 80
        .URL = "http://software.xfx.net/utilities/popupkiller/getpopups.php?FromPopUp=" & lvTitles.ListItems.Count
        .Execute
    End With
    
    Do
        DoEvents
    Loop Until StartMerging
        
    Cancel = False
    On Error Resume Next
    Do
        If Not DoneDownload Then
            vtData = inetGet.GetChunk(128, icString)
            strData = strData & vtData
            lblAction.Caption = "Downloading and Merging: " & Int(Len(strData) / 1024) & "KB"
        End If
        If vtData = "" And Not DoneDownload Then
            strData = Left$(strData, Len(strData) - 7)
            lblAction.Caption = "Merging..."
            DoneDownload = True
        End If
        
        PopUp = Split(strData, "%0D%0A%FF")(i)
        If Err.Number = 0 Then
            PopUp = URLDecode(PopUp)
            If Len(PopUp) Then
                If Asc(Right(PopUp, 1)) = 0 Then PopUp = Left(PopUp, Len(PopUp) - 1)
                If IsInList(lvTitles, PopUp, True) Then
                    e = e + 1
                Else
                    If IsInList(lvBanned, PopUp, True) Then
                        e = e + 1
                    Else
                        AddSelected PopUp, True, True
                        j = j + 1
                    End If
                End If
            End If
            lblDownloaded.Caption = "Downloaded: " & i
            lblAdded.Caption = "Added: " & j & "   (" & Int(j / (j + e) * 100) & "%)"
            lblSkipped.Caption = "Skipped: " & e & "   (" & Int(e / (j + e) * 100) & "%)"
            i = i + 1
        Else
            If DoneDownload Then
                Exit Do
            Else
                Err.Clear
            End If
        End If
        DoEvents
    Loop Until ErrorConnecting Or Cancel
    inetGet.Cancel
    
    cmdCancel.Enabled = False
    chkDownload.Value = vbUnchecked
    
    lblAction.Caption = "Saving PopUps..."
    SavePopUps INIFile
    
    MousePointer = ccDefault
    
End Sub

Private Sub DownloadBanned()

    On Error Resume Next

    Dim i As Integer
    Dim ff As Integer
    
    If Not DisableBanned Then Exit Sub
    
    If Dir(BannedFile) <> "" Then Kill BannedFile
    
    Cancel = False
    GettingBanned = True
    ErrorConnecting = False
    With inetCtrl
        .URL = "ftp://software.xfx.net"
        .UserName = "puk_user"
        .Password = "the_password"
        .RequestTimeout = 15
        If GetSetting(App.EXEName, "Proxy", "Use", 0) = 1 Then
            .AccessType = icNamedProxy
            .Proxy = GetSetting(App.EXEName, "Proxy", "Address") + ":" + GetSetting(App.EXEName, "Proxy", "Port")
        Else
            .AccessType = icUseDefault
        End If
        Err.Clear
        .Execute , "GET banned.ini """ + BannedFile + """"
        Do
            DoEvents
        Loop Until Not .StillExecuting _
                    Or ErrorConnecting _
                    Or Err.Number <> 0 _
                    Or Cancel
                    
        If Cancel Then Exit Sub
        
        If ErrorConnecting Then
            MsgBox "Error: " & .ResponseCode & vbCrLf + .ResponseInfo, vbOKOnly + vbCritical, "Error"
            Exit Sub
        ElseIf Err.Number > 0 Then
            MsgBox "Error: " & Err.Number & vbCrLf + Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            lblAction.Caption = "Loading Banned List..."
        End If
        .Execute , "CLOSE"
    End With
    
    GetBannedPopUps
    GettingBanned = False
    
End Sub

Private Function URLDecode(str As String) As String

    Dim i As Long
    Dim m As String
    
    i = InStr(1, str, "+")
    Do While i <> 0
        str = Left$(str, i - 1) + " " + Mid$(str, i + 1)
        i = InStr(i + 1, str, "+")
    Loop
    
    i = InStr(1, str, "%")
    Do While i <> 0
        Select Case Mid$(str, i + 1, 1)
            Case "%"
                str = Left$(str, i) + Mid$(str, i + 1)
                i = i + 1
            Case Else
                m = Chr(Hex2Dec(Mid$(str, i + 1, 2)))
                If m = vbCr Or m = vbLf Then m = ""
                str = Left$(str, i - 1) + m + Mid$(str, i + 3)
        End Select
        i = InStr(i, str, "%")
    Loop
    
    URLDecode = str

End Function

Function Hex2Dec(h As String) As Long

    Hex2Dec = CLng("&H" & h)

End Function

Private Sub inetCtrl_StateChanged(ByVal State As Integer)

    Select Case State
        Case icResolvingHost
            lblAction.Caption = "Resolving Host..."
        Case icConnecting
            lblAction.Caption = "Connecting..."
        Case icReceivingResponse
            If Not LoggingOff Then
                If GettingBanned Then
                    lblAction.Caption = "Getting Banned PopUps..."
                Else
                    lblAction.Caption = "Uploading PopUps..."
                End If
            End If
        Case icResponseCompleted
            lblAction.Caption = "Operation succesful, Logging Off..."
            LoggingOff = True
        Case icError
            lblAction.Caption = "ERROR: " + inetCtrl.ResponseInfo
            ErrorConnecting = True
    End Select
    
End Sub

Private Sub inetGet_StateChanged(ByVal State As Integer)

    On Error GoTo RaiseErr

    Select Case State
        Case icResolvingHost
            lblAction.Caption = "Resolving Host..."
        Case icConnecting
            lblAction.Caption = "Connecting..."
        Case icError
            ErrorConnecting = True
        Case icResponseCompleted
            StartMerging = True
        Case icError
            lblAction.Caption = "ERROR: " + inetCtrl.ResponseInfo
            Cancel = True
    End Select
    
    Exit Sub
    
RaiseErr:
    lblAction.Caption = "ERROR: " + inetCtrl.ResponseInfo
    Cancel = True

End Sub
