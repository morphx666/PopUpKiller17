Attribute VB_Name = "modIsOnLine"
Option Explicit

' Remote Access Services (RAS) APIs.
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412

Public Type RasEntryName
       dwSize As Long
       szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Type RasConn
       dwSize As Long
       hRasConn As Long
       szEntryName(RAS_MAXENTRYNAME) As Byte
       szDeviceType(RAS_MAXDEVICETYPE) As Byte
       szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" _
    (lpRasConn As Any, _
    lpcb As Long, _
    lpcConnections As Long) _
    As Long

Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" _
    (ByVal hRasConn As Long) _
    As Long
    
Declare Function nRegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) _
    As Long

Declare Function nRegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Declare Function nRegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long

Public gstrISPName As String
Public ReturnCode As Long

Public Function IsOnLine() As Boolean

    Dim hKey As Long
    Dim lpSubKey As String
    Dim phkResult As Long
    Dim lpValueName As String
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As Long
    Dim lpcbData As Long
    
    IsOnLine = -QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\RemoteAccess", "Remote Connection")
    
    Exit Function
    
    IsOnLine = False
    
    lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
    ReturnCode = nRegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)
    
    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = "Remote Connection"
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        
        ReturnCode = nRegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
        lpcbData = Len(lpData)
        ReturnCode = nRegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)
        If ReturnCode = ERROR_SUCCESS Then
            If lpData = 0 Then
                IsOnLine = False
            Else
                IsOnLine = True
            End If
        End If
    End If
    
    nRegCloseKey hKey
    
End Function

