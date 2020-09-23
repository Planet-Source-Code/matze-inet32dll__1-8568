VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Type HttpAuth definied in basGlobal


'muss noch wo anders hin implementiert werden
Public Enum ProtocolStates
    '1xx
    inetInformational = 100
    '2xx
    inetOK = 200
    inetCreated = 201
    inetAccepted = 202
    inetNoContent = 204
    '3xx
    inetMultipleChoices = 300
    inetMovedPermanently = 301
    inetMovedTemporarily = 302
    inetNotModified = 304
    '4xx
    inetBadRequest = 400
    inetUnauthorized = 401
    inetForbidden = 403
    inetNotFound = 404
    inetProxyUnauthorized = 407
    '5xx
    inetInternalServerError = 500
    inetNotImplemented = 501
    inetBadGateway = 502
    inetServiceUnavaible = 503
    
    'definied by app
    'inetCanceled = 600
    'inetTimedout = 601
    'inetNetworkError = 602
End Enum




'Data Members
Private m_HttpVersion As String
Private m_Status As ProtocolStates
Private m_Message As String
Private m_Header As String
Private m_Resource As InetURL


'************************************************
'Implementierung der Klasse

'HTTP Status Header
Public Property Get Protocol() As String: Protocol = m_HttpVersion: End Property
Public Property Get Status() As ProtocolStates: Status = m_Status: End Property
Friend Property Let Status(NewStatus As ProtocolStates): m_Status = NewStatus: End Property
Public Property Get Message() As String: Message = m_Message: End Property
Friend Property Let Message(NewMessage As String): m_Message = NewMessage: End Property

'Net Times
Public Property Get Request() As String: Request = GetValue("Date"): End Property
Public Property Get Modified() As String: Modified = GetValue("Last-Modified"): End Property

'Electronic Tag
Public Property Get ETag() As String: ETag = GetValue("ETag"): End Property

'Servernames
Public Property Get Proxy() As String: Proxy = GetValue("Via"): End Property
Public Property Get Server() As String: Server = GetValue("Server"): End Property

'Connection & Location
Friend Property Get Redirection() As String: Redirection = GetValue("Location"): End Property
Friend Property Get Connection() As String: Connection = GetValue("Connection"): End Property
Friend Property Get ProxyConnection() As String: ProxyConnection = GetValue("Proxy-Connection"): End Property

'Content
Public Property Get DocType() As String: DocType = GetValue("Content-Type"): End Property
Public Property Get Length() As Long: Length = Val(GetValue("Content-Length", , -1)): End Property


Public Property Get Resource() As InetURL
    Dim retObj As New InetURL
    If TypeName(m_Resource) = "Nothing" Then
        retObj = GetValue("Content-Location", , "")
    Else
        retObj = GetValue("Content-Location", , m_Resource.Url)
    End If
    Set Resource = retObj
End Property

Friend Property Set Resource(NewResource As InetURL)
    Set m_Resource = NewResource
End Property


'Proxy Autorisierung
Friend Property Get ProxyAuthenticate(Optional Index As Integer) As HttpAuth
    'Variablen deklaration
    Dim retAuth As HttpAuth
    Dim tmp As String
    
    'Liste erstellen
    tmp = GetValue("Proxy-Authenticate", Index)
    If tmp <> "" Then
        If lPos(tmp, " ") <> 0 Then
            retAuth.Package = Left(tmp, lPos(tmp, " ") - 1)
            retAuth.Data = Trim(Mid(tmp, lPos(tmp, " ") + 1))
        Else
            retAuth.Package = tmp
            retAuth.Data = ""
        End If
        ProxyAuthenticate = retAuth
    End If
End Property


'Www Autorisierung
Friend Property Get Authenticate(Optional Index As Integer) As HttpAuth
    'Variablen deklaration
    Dim retAuth As HttpAuth
    Dim tmp As String
    
    'Liste erstellen
    tmp = GetValue("WWW-Authenticate", Index)
    If tmp <> "" Then
        If lPos(tmp, " ") <> 0 Then
            retAuth.Package = Left(tmp, lPos(tmp, " ") - 1)
            retAuth.Data = Trim(Mid(tmp, lPos(tmp, " ") + 1))
        Else
            retAuth.Package = tmp
            retAuth.Data = ""
        End If
        Authenticate = retAuth
    End If
End Property


'Liefert den zugehörigen Value Wert zum Key
Private Function GetValue(Key As String, Optional Index As Integer, Optional Default As String) As String
    'Variablen deklaration
    Dim Headers() As String
    Dim Header As Variant
    Dim HeaderKey As String
    Dim HeaderValue As String
    Dim cIndex As Integer
    
    'Header in ein Array aufteilen
    Headers = Split(m_Header, vbCrLf)
    
    'Defaultwert für Rückgabe setzen, wenn die
    'Headerzeile nicht gefunden wird
    GetValue = Default
    
    'Headerzeile suchen
    For Each Header In Headers
        HeaderKey = Trim(Left(Header, lPos(Header, ":") - 1))
        HeaderValue = Trim(Mid(Header, lPos(Header, ":") + 1))
        
        If LCase(HeaderKey) = LCase(Key) Then
            If Index = cIndex Then
                GetValue = HeaderValue
                Exit For
            Else
                cIndex = cIndex + 1
            End If
        End If
    Next Header
End Function


'HTTP Header Parsen
Friend Sub ParseHeader(HttpHeader As String, Optional Resource As Variant)
    'Variablen deklaration
    Dim Header As String
    Header = HttpHeader
    
    'Default Resource setzen
    If TypeName(Resource) = "String" Then
        Set m_Resource = New InetURL
        m_Resource = Resource
    ElseIf TypeName(Resource) = "InetURL" Then
        Set m_Resource = Resource
    Else
        Err.Raise 13  'Typen unverträglich
    End If
    
    'vbCrLf am Ende des Headers entfernen
    Do While Right(Header, 2) = vbCrLf
        Header = Left(Header, Len(Header) - 2)
    Loop
        
    'Prüfen ob der HTTP Header gültig ist
    If UCase(Left(Header, 4)) <> "HTTP" Then
        Err.Raise httpInvalidHeader, "HttpRequest.HttpHeader", "Es wurde keine gültige Http Statuszeile zurückgeliefert."
        Exit Sub
    End If
    
    'Status Informationen und Protokoll Version ermitteln
    m_HttpVersion = Left(Header, 8)
    m_Status = Mid(Header, 10, 3)
    If lPos(Header, vbCrLf) - 14 > 0 Then
        m_Message = Left(Mid(Header, 14), lPos(Header, vbCrLf) - 14)
    Else
        m_Message = ""
    End If
    
    'Header ohne Status Informationen speichern
    m_Header = Mid(Header, lPos(Header, vbCrLf) + 2)
End Sub
