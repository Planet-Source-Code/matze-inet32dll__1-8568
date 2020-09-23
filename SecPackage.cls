VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API Declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal source As Any, ByVal Length As Long)
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpTarget As Any, ByVal lpSource As Any, ByVal iMaxLength As Long) As Long

'      ULONG fCapabilities;  // capability of bit mask
'      USHORT wVersion;      // version of driver
'      USHORT wRPCID;        // identifier for RPC run time
'      ULONG cbMaxToken;     // size of authentication token
'      SEC_CHAR * Name;      // text name
'      SEC_CHAR * Comment;   // comment

'Package Information
Private m_Capabilities As Long
Private m_Version As Integer
Private m_RPCID As Integer
Private m_MaxToken As Long
Private m_Package As String
Private m_Comment As String

'GET Methoden
Public Property Get Capabilities() As Long: Capabilities = m_Capabilities: End Property
Attribute Capabilities.VB_UserMemId = 0
Public Property Get Version() As Integer: Version = m_Version: End Property
Public Property Get RPCID() As Integer: RPCID = m_RPCID: End Property
Public Property Get MaxToken() As Long: MaxToken = m_MaxToken: End Property
Public Property Get Package() As String: Package = m_Package: End Property
Public Property Get Comment() As String: Comment = m_Comment: End Property

Private Sub InitBasicPackageValues()
    'Integriertes Basic Packages benutzen
    m_Capabilities = 0
    m_Version = 1
    m_RPCID = 0
    m_MaxToken = 512
    m_Package = "BASIC"
    m_Comment = "Integriertes Basic Package 1.0"
End Sub

'Zuweisung der Package Daten über Package Namen
Public Property Let Package(NewPackage As String)
    'Variablen deklaration
    Dim pPack As Long
    Dim Pack As SecPkgInfo
    Dim i As Long
    
    If UCase(NewPackage) = "BASIC" Then
        'Integriertes Basic Packages benutzen
        InitBasicPackageValues
    Else
        'Package Informationen anfordern
        i = QuerySecurityPackageInfo(NewPackage, pPack)
        If i = SEC_E_OK Then
            'Struktur kopieren und Speicher freigeben
            CopyMemory Pack, pPack, Len(Pack)
            i = FreeContextBuffer(pPack)
            If i <> SEC_E_OK Then
                Err.Raise i, "Inet32.SecPackage", "Fehler beim freigeben von Speicher"
            Else
                'Package Informationen über die Package Eigenschaft in die Klasse aufnehmen
                SecPkgInfo = Pack
            End If
        Else
            Err.Raise vbObjectError, "Inet32.SecPackage", "Das angegebende Security Package ist nicht gefunden wurden"
        End If
    End If
End Property

'Zuweisung der Package Daten über Struktur
Friend Property Let SecPkgInfo(PkgData As SecPkgInfo)
    With PkgData
        'Values zuweisen
        m_Capabilities = .fCapabilities
        m_Version = .wVersion
        m_RPCID = .wRPCID
        m_MaxToken = .cbMaxToken
        
        'Strings zuweisen
        m_Package = Space(256)
        lstrcpyn m_Package, .Name, 256
        m_Package = Left(m_Package, lPos(m_Package, vbNullChar) - 1)
        m_Comment = Space(256)
        lstrcpyn m_Comment, .Comment, 256
        m_Comment = Left(m_Comment, lPos(m_Comment, vbNullChar) - 1)
    End With
End Property

'Konstruktion
Private Sub Class_Initialize()
    'Standard Werte für BASIC Package laden
    InitBasicPackageValues
End Sub
