Attribute VB_Name = "basGlobal"
Public Const VER_PLATFORM_WIN32_NT = 2


'FTP Port Settings
Public Const FtpEnginePortRangeMin = 32001
Public Const FtpEnginePortRangeMax = 32099
Public FtpEnginePort As Integer


'Konstanten: InetEngine
Public Const inetEngineBase = vbObjectError + 250
  Public Const inetNotSupportedProtocol = inetEngineBase + 1

'Konstanten: InetBuffer
'Public Const inetBufferBase = vbObjectError + 200
'  Public Const inetProtected = inetBufferBase + 1

'Konstanten: HttpHeader
Public Const httpHeaderBase = vbObjectError + 50
  Public Const httpInvalidHeader = httpHeaderBase + 1

'Konstanten: HttpSecurity
Public Const secNotInitialized = vbObjectError + 100
Public Const secInvalidData = vbObjectError + 101
Public Const secFailed = vbObjectError + 102

'Konstanten: HttpEngine
Public Const httpBlocking = vbObjectError + 151
Public Const httpInvalidLink = vbObjectError + 152

'Public Type HttpAuth (für Autorisierung mit HTTP Servern)
Public Type HttpAuth
    Package As String
    Data As String
End Type



'API Declarations
Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal source As Any, ByVal Length As Long)
Public Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpTarget As Any, ByVal lpSource As Any, ByVal iMaxLength As Long) As Long

'Private Types
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


'Ermittelt die Windows Plattform
Public Function IsNT() As Boolean
    'Variablen deklaration
    Static Proceed As Boolean
    Static osvi As OSVERSIONINFO
    If Not Proceed Then
        'Plattform testen
        Dim i As Long
        osvi.dwOSVersionInfoSize = 148
        i = GetVersionEx(osvi)
    End If
    IsNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

