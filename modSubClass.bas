Attribute VB_Name = "modSubClass"
Option Explicit

Public defWindowProc2 As Long

Function WindowProc(ByVal hwnd As Long, _
                    ByVal Msg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long

    On Error Resume Next
    
    Dim retval As Long
    
    Select Case hwnd
        Case frmMain.hwnd
            Select Case Msg
                Case Else
                   retval = CallWindowProc(defWindowProc2, _
                                            hwnd, _
                                            Msg, _
                                            wParam, _
                                            lParam)
            End Select
    End Select
    
    WindowProc = retval
        
End Function

Function HiWord(dw As Long) As Integer

    If dw And &H80000000 Then
          HiWord = (dw \ 65535) - 1
    Else: HiWord = dw \ 65535
    End If
    
End Function

Function LoWord(dw As Long) As Integer

    If dw And &H8000& Then
          LoWord = &H8000 Or (dw And &H7FFF&)
    Else: LoWord = dw And &HFFFF&
    End If
    
End Function

