Attribute VB_Name = "modGlobals"
Option Explicit

Global Const intAppname As String = "PopUpKiller"

Public Enum SafeSmartConstants
    [Invalid]
    [CloseIt]
    [IgnoreIt]
    [CloseIt_and_Add2BlackList]
    [IgnoreIt_and_Add2IgnoreList]
End Enum

Public Type PreferencesProp
    AutoStart As Boolean
    JumpToHomePage As Boolean
    HomePage As String
    Smart As Boolean
    SmartSensitivity As Long
    SafeMode As Boolean
    Limit As Boolean
    LimitNum As Integer
    UseCustomIETitle As Boolean
    CustomIETitle As String
    FlashIcon As Boolean
    ShowIcon As Boolean
    ShowTabsText As Boolean
    ScrambleText As Boolean
    SoundFX As Boolean
    Changed As Boolean
    FGeo As Boolean
    DisableShortcutKeys As Boolean
    DisableWildcards As Boolean
End Type

Public FontCharSet As Long

Public Preferences As PreferencesProp
Public SafeSmartAns As SafeSmartConstants
Public SafeSmartHwnd As Long
#If DEBUGGEOPLUGIN = 0 Then
    Public FGeoLib As Object
#Else
    Public FGeoLib As New FGeo.clsFGeo
#End If

Public Function IsInList(lv As ListView, Title As String, UseWildCards As Boolean) As Boolean

    Dim blItem As ListItem
    Dim mj As String
    Dim sj As String
    Dim k() As String
    Dim i As Integer
    Dim j As Integer
    Dim si As String
    Dim WildCardMatch As Boolean
    Dim chkChar As String
    
    chkChar = Chr(255)

    If Not UseWildCards Then
        Set blItem = lv.FindItem(Title, lvwText, , lvwWhole)
        If Not blItem Is Nothing Then
            blItem.Selected = True
            blItem.EnsureVisible
            IsInList = True
            Exit Function
        End If
    Else
        si = LCase(Title)
        For Each blItem In lv.ListItems
            If blItem.Checked Then
                sj = LCase(blItem.Text)
                If InStr(sj, "*") Then
                    j = 1
                    k = Split(sj, "*")
                    For i = 0 To UBound(k)
                        If i > 0 Then If k(i - 1) <> chkChar Then Exit For
                        If k(i) = "" Then
                            k(i) = Chr(255)
                            i = i + 1
                            If i <= UBound(k) Then
                                For j = j To Len(si)
                                    If Mid(si, j, Len(k(i))) = k(i) Then
                                        j = j + Len(k(i))
                                        k(i) = chkChar
                                        Exit For
                                    End If
                                Next j
                            End If
                        Else
                            If i = 0 Then
                                If Mid(si, j, Len(k(i))) = k(i) Then
                                    j = j + Len(k(i))
                                    k(i) = chkChar
                                End If
                            Else
                                For j = j To Len(si)
                                    If Mid(si, j, Len(k(i))) = k(i) Then
                                        j = j + Len(k(i))
                                        k(i) = chkChar
                                        Exit For
                                    End If
                                Next j
                            End If
                        End If
                    Next i
                    WildCardMatch = True
                    For i = 0 To UBound(k)
                        WildCardMatch = WildCardMatch And (k(i) = chkChar)
                    Next i
                    If WildCardMatch Then
                        blItem.Selected = True
                        blItem.EnsureVisible
                        IsInList = True
                        Exit Function
                    End If
                Else
                    If si = sj Then
                        IsInList = True
                        Exit Function
                    End If
                End If
            End If
        Next blItem
    End If

End Function

Public Sub SkinMe(frm As Object, State As Boolean)
    
    Dim ctrl As Control

    For Each ctrl In frm.Controls
        If TypeOf ctrl Is CommandButton Then
            Select Case State
                Case True
                    EdgeSubClass ctrl.hwnd, sedSunkenOuter
                Case False
                    EdgeUnSubClass ctrl.hwnd
            End Select
        End If
    Next ctrl

End Sub

Public Sub SetupCharset(rForm As Form)

    Dim ctrl As Control
    Dim FontObj As Object
    
    On Error Resume Next
    
    'Exit Sub
    
    For Each ctrl In rForm.Controls
        Err.Clear
        Set FontObj = ctrl.Font
        If Err.Number = 0 Then
            FontObj.Charset = FontCharSet
        End If
    Next ctrl

End Sub

