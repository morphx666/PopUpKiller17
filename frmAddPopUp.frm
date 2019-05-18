VERSION 5.00
Begin VB.Form frmAddPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add PopUp"
   ClientHeight    =   1905
   ClientLeft      =   7770
   ClientTop       =   6345
   ClientWidth     =   3720
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
   ScaleHeight     =   1905
   ScaleWidth      =   3720
   Begin VB.OptionButton opList 
      Caption         =   "Add to the Exclusions List"
      Height          =   225
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   1095
      Width           =   3270
   End
   Begin VB.OptionButton opList 
      Caption         =   "Add to the Black List"
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   870
      Width           =   3270
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2595
      TabIndex        =   3
      Top             =   1500
      Width           =   1035
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   360
      Left            =   1470
      TabIndex        =   2
      Top             =   1500
      Width           =   1035
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PopUp Title"
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   960
   End
End
Attribute VB_Name = "frmAddPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim fItm As ListItem

    With frmMain
        If opList(0).Value Then
            If IsInList(.lvTitles, txtLabel.Text, True) Then
                MsgBox "The Selected PopUp cannot be added because already exists in the Black List", vbInformation + vbOKCancel, "Unable to add popup"
            Else
                If IsInList(.lvBanned, txtLabel.Text, True) Then
                    MsgBox "The Selected PopUp cannot be added because is a banned popup", vbInformation + vbOKCancel, "Unable to add popup"
                Else
                    Set fItm = .lvExclude.FindItem(txtLabel.Text, lvwText, , lvwWhole)
                    If Not fItm Is Nothing Then
                        If MsgBox("The PopUp " + txtLabel.Text + " is already listed in the Exclusion list. Are you sure you want to add it to the black list and remove it from the Exclusion list?", vbQuestion + vbYesNo, "Confirm Add PopUp") = vbYes Then
                            .AddSelected txtLabel.Text, True, False
                            .lvExclude.ListItems.Remove fItm.Index
                        End If
                    Else
                        .AddSelected txtLabel.Text, True, False
                    End If
                End If
            End If
        Else
            If Not IsInList(.lvExclude, txtLabel.Text, True) Then
                If IsInList(.lvBanned, txtLabel.Text, True) Then
                    MsgBox "The Selected PopUp cannot be added because is a banned popup", vbInformation + vbOKCancel, "Unable to add popup"
                Else
                    Set fItm = .lvTitles.FindItem(txtLabel.Text, lvwText, , lvwWhole)
                    If Not fItm Is Nothing Then
                        If MsgBox("The PopUp " + txtLabel.Text + " is already listed in the black list. Are you sure you want to add it to the Exclusion list and remove it from the black list?", vbQuestion + vbYesNo, "Confirm Add PopUp") = vbYes Then
                            .lvExclude.ListItems.Add , , txtLabel.Text
                            .lvTitles.ListItems.Remove fItm.Index
                        End If
                    Else
                        .lvExclude.ListItems.Add , , txtLabel.Text
                    End If
                End If
            End If
        End If
    End With
    
    Unload Me

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Move Screen.Width / 2 - Width / 2, Screen.Height / 2 - Height / 2
    SetupCharset Me

End Sub
