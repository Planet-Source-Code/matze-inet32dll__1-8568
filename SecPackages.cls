VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Package Information
Private m_Packages As Collection

'API Declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal source As Any, ByVal length As Long)

'Liefert die Anzahl der Security Packages zurück
Public Property Get Count() As Integer
    Count = m_Packages.Count
End Property

'Liefert die Security Packages Objekte zurück
Public Property Get Package(Index As Variant) As SecPackage
Attribute Package.VB_UserMemId = 0
    'Variablen deklaration
    Dim Pack As SecPackage
    
    If TypeName(Index) = "String" Then
        'Package nach Namen suchen
        For Each Pack In m_Packages
            If UCase(Pack.Package) = UCase(Index) Then
                Set Package = Pack
                Exit For
            End If
        Next Pack
    Else
        'Package nach Index Nummer suchen
        Set Package = m_Packages(Index + 1)
    End If
End Property


Private Sub Class_Initialize()
    'Variablen deklaration
    Dim tmpPack As New SecPackage
    Dim pPack As Long
    Dim Pack() As SecPkgInfo
    Dim nCount As Long
    Dim nRet As Long
    Dim i As Long
    
    'Basic Package initialisieren
    tmpPack.Package = "BASIC"
    
    'Package Collection initialisiren
    Set m_Packages = New Collection
    m_Packages.Add tmpPack
    
    'Security Service Provider Packages initialisieren
    nRet = EnumerateSecurityPackages(nCount, pPack)
    If nRet = SEC_E_OK Then
        ReDim Pack(nCount - 1)
        CopyMemory Pack(0), pPack, Len(Pack(0)) * nCount
        
        For i = LBound(Pack) To UBound(Pack)
            Set tmpPack = New SecPackage
            tmpPack.SecPkgInfo = Pack(i)
            m_Packages.Add tmpPack
        Next i
        
        nRet = FreeContextBuffer(pPack)
        If nRet <> SEC_E_OK Then
            Err.Raise nRet, "Inet32.SecPackages", GetMsg(nRet)
        End If
    Else
        Err.Raise nRet, "Inet32.SecPackages", GetMsg(nRet)
    End If
End Sub
