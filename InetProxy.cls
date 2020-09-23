VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InetProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Constants
Private Const inetRegPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"

'Public Enums
Public Enum AccessConstants
    inetDirect = 0
    inetNamedProxy = 1
    inetDefault = 2
End Enum

'Private Types
Private Type ProxySplit
    Server As String
    Port As Integer
End Type

'Data Members
Private m_Protocol As String
Private m_Access As AccessConstants
Private m_Override As String

'Proxy Server Adressen
Private m_Server As String
Private m_Port As Integer

'Proxy Authorisierungs Daten
Private m_Username As String
Private m_Password As String
Private m_Domain As String

'GET Methods
Public Property Get Access() As AccessConstants: Access = m_Access: End Property
Public Property Get Protocol() As String: Protocol = m_Protocol: End Property

'Servername zurückgeben
Public Property Get Server() As String
    'Variablen deklaration
    Dim Reg As New Registry
    
    If m_Access = inetDefault Then
        Reg.OpenRegistry HKEY_CURRENT_USER, inetRegPath
        Server = GetProxy(Reg.GetValue("ProxyServer"), m_Protocol).Server
        Reg.CloseRegistry
    Else
        Server = m_Server
    End If
End Property


'Liefert den Port des Proxy Server zurück
Public Property Get Port() As Integer
    'Variablen deklaration
    Dim Reg As New Registry
    
    If m_Access = inetDefault Then
        Reg.OpenRegistry HKEY_CURRENT_USER, inetRegPath
        Port = GetProxy(Reg.GetValue("ProxyServer"), m_Protocol).Port
        Reg.CloseRegistry
    Else
        Port = m_Port
    End If
End Property


'Liefert die Ausnahmen zurück, bei denen der Proxy nicht verwendet wird
Public Property Get Override() As String
    'Variablen deklaration
    Dim Reg As New Registry

    If m_Access = inetDefault Then
        Reg.OpenRegistry HKEY_CURRENT_USER, inetRegPath
        Override = Reg.GetValue("ProxyOverride")
        Reg.CloseRegistry
    Else
        Override = m_Override
    End If
End Property


'Benutzerautorisierung
Public Property Get Username() As String: Username = m_Username: End Property
Public Property Get Password() As String: Password = m_Password: End Property
Public Property Get Domain() As String: Domain = m_Domain: End Property


'Setzt eine neue Zugriffsart auf den Proxy
Public Property Let Access(NewAccess As AccessConstants)
    m_Access = NewAccess
End Property


'Setzt das Protokoll für den der Server verwendet werden soll
Public Property Let Protocol(NewProtocol As String)
    m_Protocol = NewProtocol
End Property


'Weist einen Server Namen für den Proxy zu
Public Property Let Server(NewProxyServer As String)
    If lPos(NewProxyServer, ":") <> 0 Then
        m_Server = Left(NewProxyServer, lPos(NewProxyServer, ":") - 1)
        m_Port = Mid(NewProxyServer, lPos(NewProxyServer, ":") + 1)
    Else
        m_Server = NewProxyServer
    End If
End Property


'Weist ein neuen Server Port zu
Public Property Let Port(NewProxyPort As Integer)
    m_Port = NewProxyPort
End Property


'Setzt die Ausnahmen in denen kein Proxy Server verwendet werden soll
Public Property Let Override(NewOverrides As String)
    m_Override = NewOverrides
End Property


'Setzt neuen Benutzernamen
Public Property Let Username(NewUsername As String)
    If lPos(NewUsername, "\") <> 0 Then
        m_Username = Mid(NewUsername, lPos(NewUsername, "\") + 1)
    Else
        m_Username = NewUsername
    End If
End Property


'Setzt neues Kennwort
Public Property Let Password(NewPassword As String)
    m_Password = NewPassword
End Property


'Setzt neue Domäne
Public Property Let Domain(NewDomain As String)
    If lPos(NewDomain, "\") <> 0 Then
        m_Domain = Left(NewDomain, lPos(NewDomain, "\") - 1)
    Else
        m_Domain = NewDomain
    End If
End Property


'Prüft ob der Proxy benutzt werden soll
Friend Function UseProxy(Resource As Variant) As Boolean
    'Variablen deklaration
    Dim ptrResource As InetURL
    Dim Reg As New Registry
    
    'Eingabe Parameter überprüfen
    If TypeName(Resource) = "InetURL" Then
        Set ptrResource = Resource
    Else
        Set ptrResource = New InetURL
        ptrResource.Url = Resource
    End If
    
    'Access Eigenschaft auswerten
    Select Case m_Access
        Case inetDirect
            'Kein Proxy Server verwenden
            UseProxy = False
            
        Case inetNamedProxy
            'Bennanten Proxy Server aus Eigenschaften nehmen
            If m_Server = "" Then
                'Kein Proxy Server eingetragen
                UseProxy = False
            Else
                'Proxy benutzten, wenn keine Ausnahme besteht
                UseProxy = Not IsOverride(ptrResource, Override)
            End If
            
        Case inetDefault
            'Proxy Server des Internet Explorers benutzen
            Reg.OpenRegistry HKEY_CURRENT_USER, inetRegPath
            If Left(Reg.GetValue("ProxyEnable"), 2) = "01" Then
                If Reg.GetValue("ProxyServer") <> "" Then
                    If GetProxy(Reg.GetValue("ProxyServer"), m_Protocol).Server <> "" Then
                        'Bis hierhin OK, wenn keine Ausnahme den Proxy wieder abschaltet
                        If IsOverride(ptrResource, Override) = False Then
                            'Alle Einstellungen korrekt Proxy Server verwenden
                            UseProxy = True
                        Else
                            UseProxy = False
                        End If
                    Else
                        UseProxy = False
                    End If
                Else
                    UseProxy = False
                End If
            Else
                UseProxy = False
            End If
            Reg.CloseRegistry
            
    End Select
End Function


'Prüft ob die durch ein Override der Proxy abgeschaltet wird
Private Function IsOverride(Resource As InetURL, Overrides As String) As Boolean
    'Variablen deklaration
    Dim arrOverrides() As String
    Dim Exception As Variant
    
    'Ausnahme suchen
    arrOverrides = Split(Overrides, ";")
    For Each Exception In arrOverrides
        Exception = Trim(Exception)   'Spaces entfernen
        If Left(Exception, 1) = "*" Then
            '*.205.3; *.brm01.telekom.de  Einträge prüfen
            If LCase(Right(Resource.Hostname, Len(Exception) - 1)) = LCase(Mid(Exception, 2)) Then
                IsOverride = True
                Exit For
            End If
        ElseIf Right(Exception, 1) = "*" Then
            '192.168.* www.*  Einträge prüfen
            If LCase(Left(Resource.Hostname, Len(Exception) - 1)) = LCase(Left(Exception, Len(Exception) - 1)) Then
                IsOverride = True
                Exit For
            End If
        ElseIf LCase(Resource.Hostname) = LCase(Exception) Then
            'www.t-online.de = www.t-online.de Einträge prüfen
            IsOverride = True
            Exit For
        End If
    Next Exception
End Function


'Analysiert einen String, in dem Proxy Server aufgezählt werden
Private Function GetProxy(Proxies As String, Optional ForProtocol As String) As ProxySplit
    'Variablen deklaration
    Dim Ptr As Variant
    Dim Server As String
    Dim Port As Integer
    Dim tmpProtocol As String
    Dim arrProxies() As String
    
    If lPos(Proxies, "=") <> 0 Then
        'Proxy Aufzählung
        arrProxies = Split(Proxies, ";")
        'Protokoll auswerten, Standard HTTP
        If ForProtocol = "" Then
            tmpProtocol = "HTTP"
        Else
            tmpProtocol = ForProtocol
        End If
        'Proxy Aufzählung durchsuchen
        For Each Ptr In arrProxies
            'Proxy für das angegebende Protokoll suchen
            Ptr = Trim(Ptr)
            If lPos(Ptr, "=") > 0 Then
                If LCase(Left(Ptr, lPos(Ptr, "=") - 1)) = LCase(tmpProtocol) Then
                    'Proxy gefunden, Daten auswerten
                    Server = Mid(Ptr, lPos(Ptr, "=") + 1)
                    Port = Mid(Server, lPos(Server, ":") + 1)
                    Server = Left(Server, lPos(Server, ":") - 1)
                    GetProxy.Server = Server
                    GetProxy.Port = Port
                    'Methode verlassen
                    Exit Function
                End If
            End If
        Next Ptr
        'Es wurde kein Proxy für das angegebende Protokoll gefunden
    Else
        'Nur ein Proxy für alles
        GetProxy.Server = Left(Proxies, lPos(Proxies, ":") - 1)
        GetProxy.Port = Mid(Proxies, lPos(Proxies, ":") + 1)
    End If
End Function
