VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'In dieser Klasse ist das HTTP Protokoll implementiert und zwar
'in den Schnittstellen: InetEngine und HttpEngine

'Debug Flag
Private Const DEBUG_MODE = False

'Ereignisse
Event Connecting(ByVal Proxy As String, ByVal Host As String)
Event Connected()
Event Requesting(ByVal Methode As String, ByVal Resource As InetURL)
Event HeaderReceived()
Event Progress(ByVal bytesReceived As Long, ByVal bytesTotal As Long)
Event Challenge()
Event Redirection(ByVal FromUrl As InetURL, ToUrl As InetURL)
Event Complete()
Event Error(ByVal Number As Long, ByVal Reason As String)

'Objekt Verweise
Private frxHttpContainer As frmHttpContainer
Private WithEvents sckHttp As Winsock
Attribute sckHttp.VB_VarHelpID = -1
Private WithEvents tmrConnection As Timer
Attribute tmrConnection.VB_VarHelpID = -1
Private WithEvents tmrData As Timer
Attribute tmrData.VB_VarHelpID = -1



'Public State Enum  (muss noch wo anders hin implementiert werden)
Enum InetStates
    inetUnknown = 0
    inetConnecting = 1
    inetConnected = 2
      
    inetExecuting = 3
    inetSending = 4
    inetReceiving = 5
    inetComplete = 6
    inetError = 9
    inetTimedout = 11
    inetCanceled = 10
End Enum


'Flag ob dieses Objekt als FTP Proxy verwendet wird
Private m_FtpProxy As Boolean


'Data Members
Private m_Proxy As New InetProxy
Private m_Timeout As New InetTimeout
Private m_Resource As New InetURL
Private m_Methode As String
Private m_SendBuffer As String
Private m_HeaderReceived As Boolean
Private m_HeaderBuffer As String
Private m_Header As New HttpHeader
Private m_Buffer As New InetBuffer
Private m_State As InetStates
Private m_Tag As Variant
Private m_Callback As Object
Private m_Blocking As Boolean
Private m_Security As HttpSecurity      'SSPI Security Interface
Private m_ProxySecurity As HttpSecurity 'SSPI Security Interface


'*
'*
'*
'*
'--------------------------------------------------------
' Implementierung der InetEngine Schnittstelle
'--------------------------------------------------------
'Implements InetEngine

'GET Methoden
'Private Property Get InetEngine_Callback() As Object: Set InetEngine_Callback = m_Callback: End Property
'Private Property Get InetEngine_Resource() As InetURL: Set InetEngine_Resource = m_Resource: End Property
'Private Property Get InetEngine_Proxy() As InetProxy: Set InetEngine_Proxy = m_Proxy: End Property
'Private Property Get InetEngine_Timeout() As InetTimeout: Set InetEngine_Timeout = m_Timeout: End Property
'Private Property Get InetEngine_Buffer() As InetBuffer: Set InetEngine_Buffer = m_Buffer: End Property
'Private Property Get InetEngine_Header() As InetHeader: Set InetEngine_Header = m_Header: End Property
'Private Property Get InetEngine_Blocking() As Boolean: InetEngine_Blocking = m_Blocking: End Property
'Private Property Get InetEngine_State() As InetStates: InetEngine_State = m_State: End Property
'Private Property Get InetEngine_Tag() As Variant: Set InetEngine_Tag = m_Tag: End Property

'SET Methoden
'Private Property Set InetEngine_Callback(ByVal CallbackObj As Object): Set m_Callback = CallbackObj: End Property
'Private Property Set InetEngine_Proxy(ByVal ProxyObj As InetProxy): Set Proxy = ProxyObj: End Property
'Private Property Set InetEngine_Resource(ByVal Url As InetURL): Set Resource = Url: End Property
'Private Property Set InetEngine_Tag(NewTag As Variant): Set m_Tag = NewTag: End Property
'Private Property Set InetEngine_Timeout(ByVal NewTimeout As InetTimeout): Set Timeout = NewTimeout: End Property

'LET Methoden
'Private Property Let InetEngine_Tag(NewTag As Variant): m_Tag = NewTag: End Property

'Öffnet eine URL und gibt den Inhalt zurück
'Private Function InetEngine_OpenURL(Resource As Variant) As InetBuffer
'    Set InetEngine_OpenURL = OpenURL(Resource, "GET")
'End Function

'Füht Operationen auf dem Server aus
'Private Sub InetEngine_Execute(Command As String, Optional Argument As Variant)
'    Stop
'End Sub

'Bricht eine Anforderung ab
'Private Sub InetEngine_Cancel()
'    Cancel
'End Sub



'FTP Proxy Flag setzen, lesen
Friend Property Get FtpProxy() As Boolean: FtpProxy = m_FtpProxy: End Property
Friend Property Let FtpProxy(State As Boolean): m_FtpProxy = State: End Property


'*
'*
'*
'*
'--------------------------------------------------------
' Implementierung der HttpEngine Schnittstelle
'--------------------------------------------------------
'*
'*
'*
'*

'GET Methoden
Public Property Get Timeout() As InetTimeout: Set Timeout = m_Timeout: End Property
Public Property Get Proxy() As InetProxy: Set Proxy = m_Proxy: End Property
Public Property Get Methode() As String: Methode = m_Methode: End Property
Public Property Get Header() As HttpHeader: Set Header = m_Header: End Property
Public Property Get Buffer() As InetBuffer: Set Buffer = m_Buffer: End Property
Public Property Get State() As InetStates: State = m_State: End Property
Public Property Get Resource() As InetURL: Set Resource = m_Resource: End Property

'Setzt einen neuen Objekt Status
Friend Property Let State(NewState As InetStates)
    m_State = NewState
End Property


'Referenziert ein neues Timeout Objekt
Public Property Set Timeout(NewTimeout As InetTimeout)
    If Not m_Timeout.Protected Then
        Set m_Timeout = NewTimeout
    Else
        Err.Raise vbObjectError, "Inet32.HttpSocket", "Diese Eigenschaft ist Schreibgeschützt."
    End If
End Property


'Referenziert ein neues Proxy Objekt
Public Property Set Proxy(NewProxy As InetProxy)
    Set m_Proxy = NewProxy
End Property


'Referenziert ein neues Url Objekt
Public Property Set Resource(Url As InetURL)
    If Not m_Blocking Then
        Set m_Resource = Url
    Else
        Err.Raise vbObjectError, "Inet32.HttpSocket", "Diese Eigenschaft ist Schreibgeschützt."
    End If
End Property


'Fordert ein Dokument von einem Web Server an
Public Function OpenURL(Resource As Variant, Optional Sync As Boolean = True, Optional Methode As String, Optional Data As String) As InetBuffer
    'Prüfen ob die Methode noch aktiv ist
    If m_Blocking Then
        Err.Raise httpBlocking, "HttpRequester.HttpRequest", "Es wurde versucht die Methode Request aufzurufen, während das Objekt ein Request noch verarbeitet"
        Exit Function
    End If
    m_Blocking = True
    
    'Eingabe Parameter analysieren
    If TypeName(Resource) = "Nothing" Then
        'Die Url wurde vorher schon durch die "Resource"
        'Eigenschaft übergeben
    ElseIf TypeName(Resource) = "InetURL" Then
        Set m_Resource = Resource
    Else
        Set m_Resource = New InetURL
        m_Resource.Url = Resource
    End If
    
    'Prüfen ob dieses Objekt zum Anfordern eines FTP oder HTTP Dokumentes ist
    If m_FtpProxy Then
        If Not m_Proxy.UseProxy(Resource) Then
            'Proxy Einstellungen verstossen gegen die Richtlinien
            'Dieses Objekt kann nur FTP Dokumente über einen Proxy
            'anfordern, für eine Direkte Verbindung ist ein FtpEngine
            'Objekt zuständig
            Err.Raise vbObjectError, "Inet32.HttpEngine", "Dieses Objekt kann nicht als FTP Proxy verwendet werden."
        End If
    Else
        'Prüfen ob es sich um ein HTTP Link handelt
        If Not UCase(m_Resource.Protocol) = "HTTP" Then
            Err.Raise httpInvalidLink, "Inet32.HttpEngine", "Diese Klassse unterstützt nur Http Links."
        End If
        'Prüfen ob es sich um eine absolute URL handelt
        If m_Resource.Hostname = "" Then
            Err.Raise httpInvalidLink, "Inet32.HttpEngine", "Sie müssen eine absolute URL übergeben"
        End If
    End If
    
    'Anforderungs Methode vervollständigen
    If Methode = "" Then
        m_Methode = "GET"
    ElseIf UCase(Methode) = "GET" Then
        m_Methode = "GET"
    ElseIf UCase(Methode) = "HEAD" Then
        m_Methode = "HEAD"
    'ElseIf UCase(Methode) = "POST" Then
    '    m_Methode = "POST"
    Else
        m_Methode = "GET"
    End If
    
    'Zusendende Daten in Members speichern
    m_SendBuffer = Data
    
    'Objekte zurücksetzen
    Set m_Header = New HttpHeader
    Set m_Header.Resource = m_Resource
    Set m_Security = Nothing
    Set m_ProxySecurity = Nothing
        
    'Mit Server verbinden
    Connect
    
    'Warten bis die Operation beendet ist
    If Sync Then
        Do While m_Blocking
            DoEvents
        Loop
    End If
    
    'Buffer Objekt zurückgeben
    Set OpenURL = m_Buffer
End Function


'Bricht eine Anforderung ab
Public Sub Cancel()
    Set m_Header = New HttpHeader
    Set m_Buffer = New InetBuffer
    Disconnect inetCanceled
    RaiseEvent Error(inetCanceled, "Abbruch durch Benutzer")
End Sub


'Prüft ob die Resource vorhanden ist
Public Function Exist(Resource As Variant) As Boolean
    OpenURL Resource, True, "HEAD"
    Select Case m_Header.Status
        Case 200
            Exist = True
        Case Else
            Exist = False
    End Select
End Function





'---------------------------------------------------------
' Private Klassen Members
'---------------------------------------------------------

'Verbindet sich mit dem Web Server bzw. dem Proxy Server
Private Sub Connect()
    'Variablen deklaration
    Dim Server As String
    Dim Port As Integer
    
    'Falls eine Verbindung besteht, Verbindung trennen
    sckHttp.Close
    Do While sckHttp.State <> sckClosed: DoEvents: Loop
    
    'Mit Server verbinden
    m_State = inetConnecting
    'RaiseEvent Connecting(m_Proxy.Server & ":" & m_Proxy.Port, m_Resource.Hostname & ":" & m_Resource.Hostport)
    If m_Proxy.UseProxy(m_Resource) Then
        RaiseEvent Connecting(m_Proxy.Server & ":" & m_Proxy.Port, m_Resource.Hostname & ":" & m_Resource.Hostport)
        sckHttp.Connect m_Proxy.Server, m_Proxy.Port
    Else
        RaiseEvent Connecting("", m_Resource.Hostname & ":" & m_Resource.Hostport)
        sckHttp.Connect m_Resource.Hostname, m_Resource.Hostport
    End If
    
    'Timeout Counter aktivieren
    tmrConnection.Interval = m_Timeout.Connection
    tmrConnection.Enabled = True
End Sub


'Verbindung zu einem Server abbrechen
Private Sub Disconnect(State As InetStates)
    'Verbindung zu Server schließen
    tmrConnection.Enabled = False
    tmrData.Enabled = False
    m_State = State
    sckHttp.Close
    
    'Objekt Verweise löschen
    Set m_Security = Nothing
    Set m_ProxySecurity = Nothing
    
    'Blocking State zurücksetzen
    m_Blocking = False
End Sub


'Sendet ein Request an einen Web Server
Private Sub SendRequest(Optional Header As String)
    'Variablen deklaration
    Dim tmpString As String
    Dim tmpResource As New InetURL
    
    
    
    If sckHttp.State = sckClosed Then Stop
    
    
    
    'Textmarke aus der Url entfernen
    tmpResource = m_Resource
    tmpResource.Textmark = ""
    
    'Callback Ereignis hier auslösen
    RaiseEvent Requesting(m_Methode, tmpResource)
    
    '****************************
    'Request erstellen
    '****************************
    If m_Proxy.UseProxy(tmpResource) Then
        'Anforderung über einen Proxy Server abschicken
        tmpString = m_Methode & " " & tmpResource & " HTTP/1.0" & vbCrLf
        If TypeName(m_ProxySecurity) <> "Nothing" Then
            tmpString = tmpString & "Proxy-Connection: Keep-Alive" & vbCrLf
            If Not m_ProxySecurity.HttpAuthorization(tmpString) Then
                'Anforderung abbrechen
                'Finish 407, "Der Proxy Autorisierung Vorgang ist fehlgeschlagen"
                RaiseEvent Error(407, "Der Proxy Autorisierung Vorgang ist fehlgeschlagen")
                Disconnect inetError
                Exit Sub
            End If
        End If
    Else
        'Anforderung direkt an Web Server
        tmpString = m_Methode & " " & tmpResource.UrlPath & " HTTP/1.0" & vbCrLf
    End If
    tmpString = tmpString & "Accept: text/html text/plain */*" & vbCrLf & _
                            "User-Agent: Inet32 Ctrl 1.0" & vbCrLf & _
                            "Host: " & tmpResource.Hostname & vbCrLf

    If TypeName(m_Security) <> "Nothing" Then
        tmpString = tmpString & "Connection: Keep-Alive" & vbCrLf
        If Not m_Security.HttpAuthorization(tmpString) Then
            'Anforderung abbrechen
            'Finish 401, "Der Autorisierung Vorgang ist fehlgeschlagen"
            RaiseEvent Error(401, "Der Autorisierung Vorgang ist fehlgeschlagen")
            Disconnect inetError
            Exit Sub
        End If
    End If

    'Content Inhalt vom letzten Request löschen
    Set m_Buffer = New InetBuffer

    'Optionalen Header und Content an den String einfügen
    If m_SendBuffer <> "" Then tmpString = tmpString & "Content-Length: " & Len(m_SendBuffer) & vbCrLf
    tmpString = tmpString & Header & vbCrLf & m_SendBuffer

    'Empfangsbuffer initialisieren
    Set m_Header = New HttpHeader
    m_HeaderReceived = False
    m_HeaderBuffer = ""
    Set m_Buffer = New InetBuffer

    '***********************
    'Request abschicken
    '***********************
    'm_State = httpSendRequest
    m_State = inetSending
    If DEBUG_MODE Then Debug.Print "Sende: " & vbCrLf & tmpString
    sckHttp.SendData tmpString
End Sub


'Interpretiert den Status Code der vom Remote Server geliefert wird
Private Sub InterpretStatusCode()
    'Variablen deklaration
    Dim tmpFrom As InetURL
       
    'Prüfen ob die Verbindung geschlossen werden soll
    If Not (LCase(m_Header.Connection) = "keep-alive" Or _
       LCase(m_Header.ProxyConnection) = "keep-alive") And _
       Not ((m_Header.Status = 401) Or (m_Header.Status = 407)) Then
        'Warten bis die Verbindung wirklich geschlossen ist
        sckHttp.Close
    End If
 
    'Statuscode auswerten
    Select Case m_Header.Status
        Case 100 To 299
            '1xx Information
            '2xx Ok
            Disconnect inetComplete
            RaiseEvent Complete

        Case 300 To 399
            '3xx Fehler Request Objekt neu anfordern location prüfen
            If m_Header.Redirection <> "" Then
                Set tmpFrom = m_Resource
                Set m_Resource = New InetURL
                'falls sich eine relative Url in .Redirection
                'befindet, absolute URL generieren
                m_Resource = tmpFrom
                m_Resource.UrlPath = m_Header.Redirection
                RaiseEvent Redirection(tmpFrom, m_Resource)
                Connect     'Erneut verbinden
            Else
                Disconnect inetComplete
                RaiseEvent Complete
            End If
        
        Case 401
            'Nicht Autorisiert
            'Request erneut erstellen, diesmal mit Authorization Header
            If TypeName(m_ProxySecurity) = "Nothing" Then
                If TypeName(m_Security) = "Nothing" Then
                    If m_Resource.Username <> "" Then
                        'Security Objekt initialisieren
                        Set m_Security = New HttpSecurity
                        m_Security.Resource = m_Resource
                        m_Security.UseForProxy = False
                        'HttpAuthenticate
                        If Not m_Security.HttpAuthenticate(m_Header) Then
                            Disconnect inetError
                            RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                        Else
                            RaiseEvent Challenge    'Ereignis auslösen
                            Connect                 'Erneut verbinden
                        End If
                    Else
                        Disconnect inetError
                        RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                    End If
                    
                Else
                    'Autorisierung fortsetzen
                    If m_Security.IsInSequence Then
                        If m_Security.HttpAuthenticate(m_Header) Then
                            SendRequest
                        Else
                            Disconnect inetError
                            RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                        End If
                    Else
                        Disconnect inetError
                        RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                    End If
                End If
            Else
                Disconnect inetError
                RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
            End If
            
        Case 407
            'Proxy Authorization failed
            'Request erneut erstellen, diesmal mit Authorization Header
            If TypeName(m_ProxySecurity) = "Nothing" Then
                If m_Proxy.Username <> "" Then
                    'Security Objekt initialisieren
                    Set m_ProxySecurity = New HttpSecurity
                    Dim tmp As New InetURL
                    tmp.Url = "http://www.proxy.de/"
                    tmp.Username = m_Proxy.Username
                    tmp.Password = m_Proxy.Password
                    tmp.Domain = m_Proxy.Domain
                    m_ProxySecurity.Resource = tmp
                    m_ProxySecurity.UseForProxy = True
                    'HttpAuthenticate
                    If Not m_ProxySecurity.HttpAuthenticate(m_Header) Then
                        Disconnect inetError
                        RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                    Else
                        RaiseEvent Challenge    'Ereignis auslösen
                        Connect                 'Erneut verbinden
                    End If
                Else
                    Disconnect inetError
                    RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                End If
            Else
                'Autorisierung fortsetzen
                If m_ProxySecurity.IsInSequence Then
                    If m_ProxySecurity.HttpAuthenticate(m_Header) Then
                        
                        SendRequest
                    Else
                        Disconnect inetError
                        RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                    End If
                Else
                    Disconnect inetError
                    RaiseEvent Error(inetError, "Security Service Provider Interface verursacht Fehler")
                End If
            End If
            
        Case 400 To 599
            '4xx Client Error
            '5xx Server Error
            Disconnect inetError
            RaiseEvent Error(inetError, m_Header.Status & " " & m_Header.Message)
                        
        Case Else
            'Sonstiger Fehler Verbindung trennen und
            'Anforderung abbrechen
            Disconnect inetError
            RaiseEvent Error(inetError, "Der Server liefert einen ungültigen Statuscode")
            
    End Select
End Sub


'Ereigniss Tritt ein, wenn der Remote Server die Verbindung trennt
Private Sub sckHttp_Close()
    'Timeout deaktivieren
    tmrConnection.Enabled = False
    tmrData.Enabled = False
    
    'Event auslösen, wenn Content mit Unbekannter Länge empfangen wurde
    If m_Header.Length < 0 Then
        If m_Blocking Then
            'Die Verbindung wird fortgesetzt
            InterpretStatusCode
        Else
            m_State = inetComplete
            'RaiseEvent Complete
            'Hier ein Callback einfügen
        End If
        
    ElseIf m_Methode = "GET" Then
        If m_Buffer.Written <> m_Header.Length Then
            'Daten wurden nicht vollständig empfangen,
            'Event auslösen
            RaiseEvent Error(inetError, "Die Verbindung wurde während der übertragung abgebrochen")
        
        Else
            'Status auf Complete setzen
            m_State = inetComplete
        End If
    End If
End Sub




'Ereigniss tritt ein wenn eine Verbindung erfolgreich aufgebaut wurde
Private Sub sckHttp_Connect()
    tmrConnection.Enabled = False
    State = inetConnected
    'Hier ein Callback durchführen
    RaiseEvent Connected
    
    tmrData.Interval = m_Timeout.Data
    tmrData.Enabled = True
    'Anforderung senden
    SendRequest
End Sub


'Es wurden Datenempfangen über das WinSock Objekt
Private Sub sckHttp_DataArrival(ByVal bytesTotal As Long)
    'Variablen deklaration
    Dim Data As String
    Dim tmp As String
        
    'Überprüfen ob Daten im Empfangsbuffer bereit stehen
    If bytesTotal > 0 Then
        'Timeout Counter zurücksetzen
        tmrData.Enabled = False
        tmrData.Enabled = True
        
        'Daten aus dem Buffer holen, und im Speicher sortieren
        sckHttp.GetData Data
        If Not m_HeaderReceived Then
            m_HeaderBuffer = m_HeaderBuffer & Data
            If lPos(m_HeaderBuffer, vbCrLf & vbCrLf) <> 0 Then
                'Der Header ist jetzt vollständig übertragen
                m_HeaderReceived = True
                m_Buffer.Append Mid(m_HeaderBuffer, lPos(m_HeaderBuffer, vbCrLf & vbCrLf) + 4)
                m_HeaderBuffer = Left(m_HeaderBuffer, lPos(m_HeaderBuffer, vbCrLf & vbCrLf) - 1)
                              
                'Header parsen
                If DEBUG_MODE Then Debug.Print "Header:" & vbCrLf & m_HeaderBuffer
                m_Header.ParseHeader m_HeaderBuffer, m_Resource
            End If
        Else
            m_Buffer.Append Data
        End If
        
        If m_HeaderReceived Then
            Select Case m_Methode
                Case "GET"  ', "POST"
                    'Warten bis alle Daten übermittelt wurden, da das
                    'ganze Dokument angefortdert wurden ist
                    If m_Header.Length >= 0 Then
                        'ContentLength Header wurde vom Server gesendet
                        
                        'Progress Event auslösen
                        RaiseEvent Progress(m_Buffer.Written, m_Header.Length)
                        'Hier ein Callback durchführen
                        
                        If m_Header.Length = m_Buffer.Written Then
                            'Content ist vollständig, Event auslösen
                            tmrData.Enabled = False
                            m_State = inetComplete
                            
                            InterpretStatusCode
                            'RaiseEvent ContentReceived(m_ContentLength)
                            'Hier ein Callback durchführen
                        End If
                    Else
                        'Progress Event auslösen
                        RaiseEvent Progress(m_Buffer.Written, -1)
                        'Callback hier auslösen
                        
                        'Warten bis die Remoteseite die Verbindung schließt
                        'Weiterer Code im Close Ereignise
                    End If
                    
                Case "HEAD"
                    'Header wurde vollständig empfangen, Status Code intepretieren
                    InterpretStatusCode
                    
            End Select
        End If
    End If
End Sub


'Ereignis: Tritt ein wenn ein Fehler auf der Netzwerk Ebene auftritt
Private Sub sckHttp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Disconnect inetError
    RaiseEvent Error(inetError, Description)
End Sub

'Header wurde übertragen
Private Sub sckHttp_SendComplete()
    m_State = inetReceiving
End Sub


'Timeout Counter für Datenübertragung zurücksetzen
Private Sub sckHttp_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    tmrData.Enabled = False
    tmrData.Enabled = True
End Sub


'Timeout Ereigniss beim Aufbau der Verbindung
Private Sub tmrConnection_Timer()
    Disconnect inetTimedout
    RaiseEvent Error(inetTimedout, "Zeitüberschreitung beim Herstellen der Verbindung")
End Sub


'Timeout Ereigniss während der Datenübertragung
Private Sub tmrData_Timer()
    Disconnect inetTimedout
    RaiseEvent Error(inetTimedout, "Zeitüberschreitung während der Verbindung")
End Sub


'Objekt initialisieren
Private Sub Class_Initialize()
    'Objekt Verweise zuordnen
    Set frxHttpContainer = New frmHttpContainer
    With frxHttpContainer
        Set sckHttp = .sckHttp
        Set tmrConnection = frxHttpContainer.tmrConnection
        Set tmrData = frxHttpContainer.tmrData
    End With
        
    'Variablen initialisieren
    'm_State = httpNotConnected
    m_State = inetUnknown
End Sub
