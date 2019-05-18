Attribute VB_Name = "netAPI"
Option Explicit

Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_GLOBALNET As Long = &H2&

Public Type NETINFOSTRUCT
    cbStructure As Long
    dwProviderVersion As Long
    dwStatus As Long
    dwCharacteristics As Long
    dwHandle As Long
    wNetType As Long
    dwPrinters As Long
    dwDrives As Long
End Type

Public Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As String
        lpRemoteName As String
        lpComment As String
        lpProvider As String
End Type

Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" _
        (ByVal hEnum As Long, _
        lpcCount As Long, _
        lpBuffer As Any, _
        lpBufferSize As Long) _
        As Long

Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" _
        (ByVal dwScope As Long, _
        ByVal dwType As Long, _
        ByVal dwUsage As Long, _
        lpNetResource As NETRESOURCE, _
        lphEnum As Long) _
        As Long

Public Declare Function WNetGetNetworkInformation Lib "mpr.dll" Alias "WNetGetNetworkInformationA" _
        (ByVal lpProvider As String, _
        ByRef lpNetInfoStruct As NETINFOSTRUCT) _
        As Long
            
Public Declare Function WNetGetProviderName Lib "mpr.dll" Alias "WNetGetProviderNameA" _
        (dwNetType As Long, _
        lpProvider As String, _
        lpBufferSize As Long) _
        As Long
        
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" _
        (lpError As Long, _
        lpErrorBuf As String, _
        nErrorBufSize As Long, _
        lpNameBuf As String, _
        nNameBufSize As Long) _
        As Long

Public Declare Function GetLastError Lib "kernel32" () As Long
