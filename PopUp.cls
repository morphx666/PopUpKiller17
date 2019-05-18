VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"ListPopUps"
Attribute VB_Ext_KEY = "Member1" ,"PopUp"
Option Explicit

Public Enum dMode
    Browser = 0
    FullText = 1
    ContentsClose = 2
    ContentsAsk = 3
End Enum

'local variable(s) to hold property value(s)
Private mvarTitle As String 'local copy
Private mvarEnabled As Boolean 'local copy
Private mvarIsChild As Boolean 'local copy
Private mvarListPopUps As ListPopUps
Private mvarhWnd As Long 'local copy
Private mvarMode As Long 'local copy

Public Property Let DetectionMode(ByVal vData As dMode)

    mvarMode = vData

End Property

Public Property Get DetectionMode() As dMode

    DetectionMode = mvarMode

End Property

Public Property Let IsChild(ByVal vData As Boolean)

    mvarIsChild = vData

End Property

Public Property Get IsChild() As Boolean

    IsChild = mvarIsChild

End Property

Public Property Let hWnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hWnd = 5
    mvarhWnd = vData
End Property

Public Property Get hWnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hWnd
    hWnd = mvarhWnd
End Property

Public Property Get ListPopUps() As ListPopUps
    If mvarListPopUps Is Nothing Then
        Set mvarListPopUps = New ListPopUps
    End If


    Set ListPopUps = mvarListPopUps
End Property

Public Property Set ListPopUps(vData As ListPopUps)
    Set mvarListPopUps = vData
End Property

Private Sub Class_Terminate()
    Set mvarListPopUps = Nothing
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Enabled = 5
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Enabled
    Enabled = mvarEnabled
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property

Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property
