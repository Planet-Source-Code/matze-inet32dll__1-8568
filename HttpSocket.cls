VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Objekt Verweise
Private WithEvents sckHttp As Winsock
Attribute sckHttp.VB_VarHelpID = -1
Private WithEvents tmrTimeout As Timer
Attribute tmrTimeout.VB_VarHelpID = -1
Private frmContainer As frmContainer

'Data Members
Private m_Methode As String
Private m_State As Long
Private m_HeaderReceived As Boolean
Private m_Header As String
Private m_ContentLength As Long
Private m_Content As New InetBuffer
Private m_Timeout As New InetTimeout

'HttpSocket Events
Event TimeoutConnection()
Event Connect()
Event HeaderReceived(ByVal bytesTotal As Long)
Event TimeoutData()
Event ContentProgress(ByVal bytesReceived As Long, ByVal bytesLeft As Long)
Event ContentReceived(ByVal bytesTotal As Long)
Event Disconnect()

'GET Methoden
Public Property Get Timeout() As InetTimeout: Set Timeout = m_Timeout: End Property
Public Property Get Methode() As String: Methode = m_Methode: End Property
Public Property Get Header() As String: Header = m_Header: End Property
Public Property Get Content() As InetBuffer: Set Content = m_Content: End Property
Public Property Get State() As Long: State = m_State: End Property

'LET Methoden
Public Property Let Timeout(NewTimeout As InetTimeout)
    If Not m_Timeout.Protected Then
        m_Timeout.Connection = NewTimeout.Connection
        m_Timeout.Data = NewTimeout.Data
    Else
        Err.Raise vbObjectError, "Inet32.HttpSocket", "Diese Eigenschaft ist Schreibgeschützt."
    End If
End Property

'SET Methoden
Public Property Set Timeout(NewTimeout As InetTimeout)
    If Not m_Timeout.Protected Then
        Set m_Timeout = NewTimeout
    Else
        Err.Raise vbObjectError, "Inet32.HttpSocket", "Diese Eigenschaft ist Schreibgeschützt."
    End If
End Property


Public Sub Connect(RemoteHost As String, Optional RemotePort As Integer)
    'Mit Web Server verbinden
    m_State = httpConnecting
    If RemotePort = 0 Then
        sckHttp.Connect RemoteHost, 80
    Else
        sckHttp.Connect RemoteHost, RemotePort
    End If
    tmrTimeout.Interval = m_Timeout.Connection
    tmrTimeout.Enabled = True
End Sub

Public Sub Disconnect()
    'Verbindung zu Server schließen
    tmrTimeout.Enabled = False
    m_State = httpNotConnected
    sckHttp.Close
    RaiseEvent Disconnect
End Sub

Public Sub SendRequest(ByVal Header As String, Optional Content As String)
    'Methode ermitteln
    m_Methode = UCase(Left(Header, lPos(Header, " ") - 1))
    
    'Header vorbereiten
    If Right(Header, 4) = vbCrLf & vbCrLf Then
        'Header ist OK
    ElseIf Right(Header, 2) = vbCrLf Then
        Header = Header & vbCrLf
    Else
        Header = Header & vbCrLf & vbCrLf
    End If
    
    'Vermerken das noch kein Header empfangen
    'wurde, Empfangsbuffer löschen
    m_HeaderReceived = False
    m_Header = ""
    m_ContentLength = -1
    Set m_Content = New InetBuffer
    
    'Request abschicken
    m_State = httpSendRequest
    sckHttp.SendData Header & Content
End Sub

Private Sub sckHttp_Close()
    'Timeout deaktivieren
    tmrTimeout.Enabled = False
    
    'Event auslösen, wenn Content mit Unbekannter Länge empfangen wurde
    If m_ContentLength < 0 Then
        RaiseEvent ContentReceived(m_Content.Length)
    ElseIf m_Methode = "GET" Then
        If m_Content.Length < m_ContentLength Then
            'Daten wurden nicht vollständig empfangen,
            'Event auslösen
            RaiseEvent ContentReceived(m_Content.Length)
        End If
    End If
    
    'Event auslösen, das die Verbindung geschlossen ist
    RaiseEvent Disconnect
End Sub

Private Sub sckHttp_Connect()
    tmrTimeout.Enabled = False       'Timeout Timer für Connection zurücksetzen
    m_State = httpConnected          'Status aktualisieren
    RaiseEvent Connect               'Connect Event auslösen
    tmrTimeout.Interval = m_Timeout.Data
    tmrTimeout.Enabled = True        'Timeout Timer für Data aktivieren
End Sub

Private Sub sckHttp_DataArrival(ByVal bytesTotal As Long)
    'Variablen deklaration
    Dim Data As String
    Dim tmp As String
        
    'Überprüfen ob Daten im Empfangsbuffer bereit stehen
    If bytesTotal > 0 Then
        'Timeout Counter zurücksetzen
        tmrTimeout.Enabled = False
        tmrTimeout.Enabled = True
        
        'Daten aus dem Buffer holen, und im Speicher sortieren
        sckHttp.GetData Data
        If Not m_HeaderReceived Then
            m_Header = m_Header & Data
            If lPos(m_Header, vbCrLf & vbCrLf) <> 0 Then
                'Der Header ist jetzt vollständig übertragen
                m_HeaderReceived = True
                m_Content.Append Mid(m_Header, lPos(m_Header, vbCrLf & vbCrLf) + 4)
                m_Header = Left(m_Header, lPos(m_Header, vbCrLf & vbCrLf) - 1)
                
                'ContentLength ermitteln
                If lPos(LCase(m_Header), "content-length") <> 0 Then
                    tmp = Mid(m_Header, lPos(LCase(m_Header), "content-length"))
                    If lPos(tmp, vbCrLf) <> 0 Then tmp = Left(tmp, lPos(tmp, vbCrLf) - 1)
                    tmp = Mid(tmp, lPos(tmp, ":") + 1)
                    tmp = Trim(tmp)
                    m_ContentLength = tmp
                Else
                    m_ContentLength = -1
                End If
                
                'Event auslösen
                RaiseEvent HeaderReceived(Len(m_Header))
            End If
        Else
            m_Content.Append Data
        End If
        
        If m_HeaderReceived Then
            Select Case m_Methode
                Case "GET"
                    'Warten bis alle Daten übermittelt wurden, da das
                    'ganze Dokument angefortdert wurden ist
                    If m_ContentLength >= 0 Then
                        'ContentLength Header wurde vom Server gesendet
                        
                        'Progress Event auslösen
                        RaiseEvent ContentProgress(m_Content.Length, m_ContentLength)
                        
                        If m_ContentLength = m_Content.Length Then
                            'Content ist vollständig, Event auslösen
                            tmrTimeout.Enabled = False
                            m_State = httpComplete
                            RaiseEvent ContentReceived(m_ContentLength)
                        End If
                    Else
                        'Progress Event auslösen
                        RaiseEvent ContentProgress(m_Content.Length, -1)
                        'Warten bis die Remoteseite die Verbindung schließt
                        'Weiterer Code im Close Ereignise
                    End If
                    
                Case "HEAD"
                    'Header wurde vollständig empfangen
            End Select
        End If
    End If
End Sub

Private Sub sckHttp_SendComplete()
    m_State = httpWait4Header
End Sub

Private Sub tmrTimeout_Timer()
    tmrTimeout.Enabled = False
    sckHttp.Close
    
    If m_State = httpConnecting Then
        m_State = httpSckTimeout
        RaiseEvent TimeoutConnection
    Else
        m_State = httpSckTimeout
        RaiseEvent TimeoutData
    End If
End Sub

Private Sub Class_Initialize()
    'Objekt Verweise zuordnen
    Set frmContainer = New frmContainer
    Set sckHttp = frmContainer.sckArray0
    Set tmrTimeout = frmContainer.tmrArray0
        
    'Variablen initialisieren
    m_State = httpNotConnected
    tmrTimeout.Enabled = False
    m_ContentLength = -1
End Sub

Private Sub Class_Terminate()
    'Objekt Verweise freigeben
    Set sckHttp = Nothing
    Set tmrTimeout = Nothing
    Set frmContainer = Nothing
End Sub
