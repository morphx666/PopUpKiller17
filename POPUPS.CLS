VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PopUp"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPopUp As PopUp


Private Sub Class_Initialize()
    'create the mPopUp object when the PopUps class is created
    Set mvarPopUp = New PopUp
End Sub


Public Property Get PopUp() As PopUp
    Set PopUp = mvarPopUp
End Property


Public Property Set PopUp(vData As PopUp)
    Set mvarPopUp = vData
End Property
Private Sub Class_Terminate()
    Set mvarPopUp = Nothing
End Sub



