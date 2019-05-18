VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim objShell As Object
Dim ScanSuccess As Boolean

Public Function Scan() As Boolean

    On Error GoTo ExitSub

    Dim objShellWindows As ShellWindows
    Dim objWebBrowser As WebBrowser_V1
    Dim objDocument As Object
    Dim i As Long
    Dim url As String
    
    Dim x As Object
    
    ScanSuccess = False
    
    Set objShellWindows = objShell.Windows
    For i = 0 To objShellWindows.Count - 1
        Set objWebBrowser = objShellWindows.Item(i)
        If Not objWebBrowser Is Nothing Then
            url = LCase(objWebBrowser.LocationURL)
            If InStr(url, "http://www.geocities.com") Or _
                InStr(url, "http://geocities.yahoo.com") Or _
                InStr(url, "http://pages.yahoo.com") Or _
                InStr(url, "http://geocities.com") Then
                Set objDocument = objWebBrowser.Document
                If objDocument.frames.length > 0 Then
                    ScanFrames objDocument
                Else
                    HideGeoShitLayers objDocument
                End If
                If ScanSuccess Then Exit For
            End If
        End If
    Next i
    
ExitSub:

    Scan = ScanSuccess
    
End Function

Private Function ScanFrames(objDocFrames As Object) As Boolean

    Dim objFrame As Object
    Dim f As Integer

    For f = 0 To objDocFrames.frames.length - 1
        Set objFrame = objDocFrames.frames.Item(f)
        If objFrame.frames.length > 0 Then
            ScanFrames objFrame.frames
        Else
            HideGeoShitLayers objFrame.Document
        End If
    Next f

End Function

Private Sub HideGeoShitLayers(objFrameDocument As Object)

    Dim l As Integer
    
    On Error Resume Next

    For l = 0 To 4
        If objFrameDocument.getElementById("yl" & l).Style.visibility <> "hidden" Then
            objFrameDocument.getElementById("yl" & l).Style.visibility = "hidden"
            ScanSuccess = True
        End If
    Next l

End Sub

Private Sub Class_Initialize()

    Set objShell = CreateObject("Shell.Application")

End Sub
