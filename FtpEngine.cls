VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtpEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const DEBUG_MODE = True

'Data Members
Private frxFtpContainer As New frmFtpContainer
Private WithEvents sckCommand As Winsock
Attribute sckCommand.VB_VarHelpID = -1
Private WithEvents sckData As Winsock
Attribute sckData.VB_VarHelpID = -1
Private m_Buffer As New InetBuffer
Private m_SendFlag As Boolean
Private m_Timeout As New InetTimeout
Private m_Proxy As New InetProxy

Private m_Format As DataFormats
Private m_Passive As Boolean


Private Const m_AnonymousPass = "mail@server.de"


Public Enum DataFormats
    dataAscii = 0
    dataBinary = 1
End Enum

'Verbindet sich mit einem Ftp Server
Public Function Connect(Hostname As String, Optional Port As Integer) As Boolean
    'Variablen deklaration
    Dim Reply As New FtpReply
    Dim TimeoutTime
    
    'Bestehende Verbindungen trennen
    sckCommand.Close
    sckData.Close
    
    'Verbindung aufbauen ...
    If Port = 0 Then
        'Standard Port 21 verwenden
        sckCommand.Connect Hostname, 21
    Else
        sckCommand.Connect Hostname, Port
    End If
    
    'Warten bis die Verbindung steht ...
    TimeoutTime = Timer + (m_Timeout.Connection / 1000)
    Do While sckCommand.State <> sckConnected And sckCommand.State <> sckError And TimeoutTime > Timer: DoEvents: Loop
    
    If sckCommand.State = sckConnected Then
        'Begrüßungs Daten empfangen
        Connect = GetReply(Reply)
    Else
        'Verbindung fehlgeschlagen
        sckCommand.Close
        Connect = False
    End If
End Function

'Trennt die Verbindung
Public Sub Quit(Optional Abort As Boolean)
    'Ausloggen
    If Abort Then
        DoExecute "ABOR"
    End If
    DoExecute "QUIT"
    
    'Bestehende Verbindungen trennen
    sckCommand.Close
    sckData.Close
End Sub

'Setzt oder Liefert die Timeout Values
Public Property Get Timeout() As InetTimeout
    Set Timeout = m_Timeout
End Property

'Setzt oder Liefert ein Proxy Objekt
Public Property Get Proxy() As InetProxy
    Set Proxy = m_Proxy
End Property

'Liefert ob die Datenübertragung Passive oder Aktiv läuft
Public Property Get Passive() As Boolean
    Passive = m_Passive
End Property

'Setzt ob die nächste Datenübertragung Passive oder Aktiv statt findet
Public Property Let Passive(NewValue As Boolean)
    m_Passive = NewValue
End Property

'Liefert den Hostname des Ftp Server zurück
Public Property Get Hostname() As String
    Hostname = sckCommand.RemoteHost
End Property

'Liefert den RemotePort des Ftp Servers zurück
Public Property Get Hostport() As Integer
    Hostport = sckCommand.RemotePort
End Property

'Liefert den Inhalt einer Resource
Public Function OpenURL(Resource, Optional Format) As InetBuffer
    'Eingabeparameter überprüfen
    Dim tmpResource As InetURL
    If TypeName(Resource) = "String" Then
        Set tmpResource = New InetURL
        tmpResource = Resource
    ElseIf TypeName(Resource) = "InetURL" Then
        Set tmpResource = Resource
    Else
        'Type unverträglich
        Err.Raise 13, "Inet32.FtpEngine"
    End If
    
    'Prüfen ob das Dokument über einen Proxy Server geholt werden soll
    If m_Proxy.UseProxy(tmpResource) Then
        'Ein HttpEngine Objekt benutzen um ein Ftp Dokument anzufordern
        Dim FtpProxy As New HttpEngine
        FtpProxy.FtpProxy = True
        Set FtpProxy.Proxy = m_Proxy
        Set OpenURL = FtpProxy.OpenURL(tmpResource, True, "GET")
    Else
        'Direkte FTP Verbindung benutzen
        Connect tmpResource.Hostname, tmpResource.Hostport
        If tmpResource.Username = "" Then
            'Anonymous Login
            Login "Anonymous", m_AnonymousPass
        Else
            'Normal Login
            Login tmpResource.Username, tmpResource.Password
        End If
        'Datei anfordern
        Dim lBuffer As New InetBuffer
        Receive tmpResource.UrlPath, lBuffer
        Set OpenURL = lBuffer
    End If
End Function

'Prüft ob 'Filename' auf dem FTP Server existiert
Public Function Exist(Resource, Code As Integer) As Boolean
    'Eingabeparameter überprüfen
    Dim tmpResource As InetURL
    Dim Reply As FtpReply
    If TypeName(Resource) = "String" Then
        Set tmpResource = New InetURL
        tmpResource = Resource
    ElseIf TypeName(Resource) = "InetURL" Then
        Set tmpResource = Resource
    Else
        'Type unverträglich
        Err.Raise 13, "Inet32.FtpEngine"
    End If
    
    'Prüfen ob das Dokument über einen Proxy Server geholt werden soll
    If m_Proxy.UseProxy(tmpResource) Then
        'Ein HttpEngine Objekt benutzen um ein Ftp Dokument anzufordern
        Dim FtpProxy As New HttpEngine
        FtpProxy.FtpProxy = True
        Set FtpProxy.Proxy = m_Proxy
        FtpProxy.OpenURL tmpResource, True, "HEAD"
        Exist = (FtpProxy.Header.Status = inetOK)
        FtpProxy.Cancel
        
    Else
        'Direkte FTP Verbindung benutzen
        If Connect(tmpResource.Hostname, tmpResource.Hostport) Then
            If tmpResource.Username = "" Then
                'Anonymous Login
                Login "Anonymous", m_AnonymousPass
            Else
                'Normal Login
                Login tmpResource.Username, tmpResource.Password
            End If
            'Datei prüfen
            If DataInit Then
                m_Buffer.Clear
                Set Reply = DoExecute("RETR " & tmpResource.UrlPath)
                'Reply auf Fehlerprüfen
                If Reply.MainState <> repPreliminary Then
                    Code = Reply.Status
                    Exist = False
                    Quit
                    Exit Function
                End If
                sckData.Close
                GetReply Reply
                Quit
                Exist = True
            End If
        Else
            'Server nicht gefunden
            Code = 0
        End If
    End If
End Function

'Sendet Daten an den Kommando Interpeter
Private Function DoExecute(Data As String) As FtpReply
    If Right(Data, 2) <> vbCrLf Then Data = Data & vbCrLf
    'Aus Debug zwecken Data ausgeben
    If DEBUG_MODE Then Debug.Print Data;
    
    'Kommando senden ...
    Dim Reply As New FtpReply
    sckCommand.SendData Data
    'Reply abwarten
    GetReply Reply
    Set DoExecute = Reply
End Function



'Initialisiert eine Daten Verbindung
Friend Function DataInit() As Boolean
    'Variablen deklaration
    Dim Arg As String
    Dim Host As String
    Dim Port As Integer
    Dim Obj As FtpReply
    
    'Falls verbunden, Verbindung trennen
    If sckData.State = sckConnected Then
        DataInit = True
        Exit Function
    End If
    sckData.Close
    
    'Aktive oder Passive Verbindung?
    If m_Passive Then
        'Passive Verbindung
        'Den Port vom Server holen
        Set Obj = DoExecute("PASV")
        On Error GoTo ErrHandle
        Arg = Mid(Obj.Message, lPos(Obj.Message, "(") + 1)
        Arg = Left(Arg, lPos(Arg, ")") - 1)
        On Error GoTo 0
        'Kommas durch Punkte ersetzen
        Do
            Arg = Left(Arg, lPos(Arg, ",") - 1) & "." & Mid(Arg, lPos(Arg, ",") + 1)
        Loop While lPos(Arg, ",") > 0
        'Host und Port parsen
        Port = Mid(Arg, rPos(Arg, ".") + 1)
        Arg = Left(Arg, rPos(Arg, ".") - 1)
        Port = Port + (Mid(Arg, rPos(Arg, ".") + 1) * 256)
        Arg = Left(Arg, rPos(Arg, ".") - 1)
        Host = Arg
        'Verbinden
        If DEBUG_MODE Then Debug.Print "data connecting: " & Host & ":" & Port
        sckData.Connect Host, Port
        DataInit = True
        
    Else
        'Aktive Verbindung
        'Den Port dem Server mitteilen
        Arg = sckCommand.LocalIP
        'Port auswählen
        If FtpEnginePort < FtpEnginePortRangeMin Then
            FtpEnginePort = Int((FtpEnginePortRangeMax - FtpEnginePortRangeMin + 1) * Rnd + FtpEnginePortRangeMin)
        Else
            FtpEnginePort = FtpEnginePort + 1
            If FtpEnginePort > FtpEnginePortRangeMax Then FtpEnginePort = FtpEnginePortRangeMin
        End If
        'Eindeutigen Port zuweisen
        sckData.Bind FtpEnginePort
        sckData.Listen
        'Punkte durch Kommas ersetzen
        Do
            Arg = Left(Arg, lPos(Arg, ".") - 1) & "," & Mid(Arg, lPos(Arg, ".") + 1)
        Loop While lPos(Arg, ".") > 0
        Arg = Arg & "," & (sckData.LocalPort And &HFF00) / &H100 & "," & (sckData.LocalPort And &HFF)
        DataInit = (DoExecute("PORT " & Arg).MainState = repComplete)
    End If
    Exit Function
    
ErrHandle:
    'Fehler während des Parsens des Hostnames und des Ports
    DataInit = False
End Function

'Trennt die Daten Verbindung
Friend Sub DataClose()
    sckData.Close
End Sub

'Sendet Daten über die Data Connection
Private Function DataSend(Data As String) As Boolean
    'flag inititalisieren, wird bei kompletter versendung der Daten
    'durch das Complete Ereignis auf true gesetzt
    m_SendFlag = False
    
    'Nur senden, wenn Verbindung besteht
    If sckData.State <> sckClosed Or sckData.State = sckError Then
        'senden
        sckData.SendData Data
    Else
        'false zurückgeben, da keine verbindung
        DataSend = False
        Exit Function
    End If
    
    'Warten bis die Daten gesendet wurden
    Do While Not m_SendFlag And Not (sckData.State = sckClosed Or sckData.State = sckError)
        DoEvents
    Loop
    DataSend = m_SendFlag
End Function

'Initialisierung
Private Sub Class_Initialize()
    'Socket Objekte initialisieren
    Set sckCommand = frxFtpContainer.sckCommand
    Set sckData = frxFtpContainer.sckData
    'Zufallsgenerator neu initialisieren
    Randomize
End Sub

'Wartet auf eine Antwort vom Server
Private Function GetReply(Reply As FtpReply) As Boolean
    'Auf Antwort warten ...
    Dim RetData As String
    Dim Buffer As String
    Dim Lastline As String
    
    Do
        'Empfangen
        On Error GoTo ErrClosed
        Do
            DoEvents
            sckCommand.GetData Buffer, vbString
            RetData = RetData & Buffer
            
            If sckCommand.State = sckClosed Or sckCommand.State = sckError Then
                'Verbindung wurde geschlossen
ErrClosed:
                GetReply = False
                Exit Function
            End If
        Loop While RetData = ""
        On Error GoTo 0
        
        'Auswerten
        If rPos(Left(RetData, Len(RetData) - 2), vbCrLf) > 0 Then
            'Die letzte Zeile extra speichern
            Lastline = Mid(RetData, rPos(Left(RetData, Len(RetData) - 2), vbCrLf) + 2)
        Else
            'Es gibt nur eine Zeile im Empfangsbuffer
            Lastline = RetData
        End If
    Loop While Mid(Lastline, 4, 1) <> " "
        
    'Zum debuggen: Empfange Daten ausgeben
    If DEBUG_MODE Then Debug.Print RetData;
    
    'FtpResponse Objekt zurückgeben
    Reply.Reply = RetData
    GetReply = True
End Function

'Ereignis: um eingehende Verbindungen anzunehmen
Private Sub sckData_ConnectionRequest(ByVal requestID As Long)
    'Socket schliessen, falls geoeffnet
    sckData.Close
    Do While sckData.State <> sckClosed
        DoEvents
    Loop
    'Verbindung annehmen
    sckData.Accept requestID
End Sub

'Ereignis: Eingehende Daten auf der Data Leitung
Private Sub sckData_DataArrival(ByVal bytesTotal As Long)
    'Variablen deklaration
    Dim Data As String
    
    If bytesTotal > 0 Then
        'Daten abholen
        sckData.GetData Data, vbString
        m_Buffer.Append Data
    End If
End Sub

'Ereignis: Tritt ein wenn Daten erfolgreich gesendet wurden
Private Sub sckData_SendComplete()
    m_SendFlag = True
End Sub








' ****************************************** '
'  Implementierung der FTP Service Methoden  '
' ****************************************** '

'Autorisiert ein Benutzer
Public Function Login(Username As String, Password As String) As Boolean
    'Username senden
    If DoExecute("USER " & Username).MainState = repIntermediate Then
        'Password senden
        Login = (DoExecute("PASS " & Password).MainState = repComplete)
    End If
End Function

'Wechselt ein Verzeichnis höher
Public Function CdUp() As Boolean
    CdUp = (DoExecute("CDUP").MainState = repComplete)
End Function

'Wechselt ein Verzeichnis
Public Function ChDir(Directory As String) As Boolean
    ChDir = (DoExecute("CWD " & Directory).MainState = repComplete)
End Function

'Ein Verzeichnis erstellen
Public Function MkDir(Directory As String) As Boolean
    MkDir = (DoExecute("MKD " & Directory).MainState = repComplete)
End Function

'Löscht ein Verzeichnis
Public Function RmDir(Directory As String) As Boolean
    RmDir = (DoExecute("RMD " & Directory).MainState = repComplete)
End Function

'Liefert das aktuelle Verzeichnis zurück
Public Function CurDir() As String
    'Variablen deklaration
    Dim Reply As FtpReply
    Dim Path As String
    'Execute Command
    Set Reply = DoExecute("PWD")
    If Reply.MainState = repComplete Then
        'Parse Directory Name
        If lPos(Reply.Message, """") > 0 Then
            Path = Mid(Reply.Message, lPos(Reply.Message, """") + 1)
            CurDir = Left(Path, lPos(Path, """") - 1)
        ElseIf lPos(Reply.Message, "'") > 0 Then
            Path = Mid(Reply.Message, lPos(Reply.Message, "'") + 1)
            CurDir = Left(Path, lPos(Path, "'") - 1)
        'Else
            'Ungültiger Pfad Bezeichner
        End If
    End If
End Function

'Liefert die Betriebsystems Kennung zurück
Public Function System() As String
    'Variablen deklaration
    Dim Reply As FtpReply
    Set Reply = DoExecute("SYST")   'Execute
    'Antwort prüfen
    If Reply.MainState = repComplete Then System = Reply.Message
End Function

'Löscht eine Datei auf dem Server
Public Function Kill(FileName As String) As Boolean
    Kill = (DoExecute("DELE " & FileName).MainState = repComplete)
End Function

'Benennt eine Datei um
Public Function Name(OldName As String, NewName As String) As Boolean
    If DoExecute("RNFR " & OldName).MainState = repIntermediate Then
        Name = (DoExecute("RNTO " & NewName).MainState = repComplete)
    End If
End Function

'Liefert eine Liste mit Dateinamen vom Server
Public Function List(Optional Directory As String) As String
    'Variablen deklaration
    Dim Reply As New FtpReply
    Dim tmp As String
   
    'Datei Liste anfordern
    If Format <> dataAscii Then Format = dataAscii
    DataInit
    m_Buffer.Clear
    Set Reply = DoExecute(Trim("LIST " & Directory))
    
    'Auf Reply warten
    Do While Reply.MainState = repPreliminary
        DoEvents
        GetReply Reply
    Loop
    
    'Liste in ein String umwandeln und zurück geben
    m_Buffer.GetData tmp
    m_Buffer.Clear
    List = tmp
End Function

'Liefert das aktuelle Datenformat zurück
Public Property Get Format() As DataFormats
    Format = m_Format
End Property

'Setzt ein neues Datenformat
Public Property Let Format(NewFormat As DataFormats)
    'Variablen deklaration
    Dim Reply As FtpReply
    'Neues Daten Format dem Server mitteilen
    Select Case NewFormat
        Case dataAscii
            Set Reply = DoExecute("TYPE A")
        Case dataBinary
            Set Reply = DoExecute("TYPE I")
        Case Else
            Err.Raise vbObjectError, "Inet32.FtpEngine", "Ungültiges Datenformat"
    End Select
    'Prüfen ob das Datenformat vom Server angenommen wurden ist
    If Reply.MainState = repComplete Then m_Format = NewFormat
End Property

'Liefert ein Array mit Dateinamen
Public Function Dir(Optional Directory As String) As String()
    'Variablen deklaration
    Dim Reply As New FtpReply
    Dim tmp As String
    
    'Verzeichnis Liste anfordern
    Format = dataAscii
    DataInit
    m_Buffer.Clear
    Set Reply = DoExecute(Trim("NLST " & Directory))
    
    'Reply abwarten
    Do While Reply.MainState = repPreliminary
        DoEvents
        GetReply Reply
    Loop
    
    'String in ein Array umwandeln und zurückgeben
    m_Buffer.GetData tmp
    m_Buffer.Clear
    Dir = Split(tmp, vbCrLf)
End Function

'Initialisiert ein Daten Transfer
'Server -> Client
Public Function Receive(ServerFile As String, LocalFile As Variant) As Boolean
    'Variablen deklaration
    Dim Reply As FtpReply
    Dim tmp As String
    Dim LocalBuffer As InetBuffer
    
    'Eingabeparameter überprüfen
    If Not (TypeName(LocalFile) = "String" Or TypeName(LocalFile) = "InetBuffer") Then
        Err.Raise 13  'Typen unverträglich
    End If
    
    'Datei empfangen
    DataInit
    m_Buffer.Clear
    Set Reply = DoExecute("RETR " & ServerFile)
    'Reply auf Fehlerprüfen
    If Reply.MainState <> repPreliminary Then
        Receive = False
        Exit Function
    End If
    
    'Warten bis die Daten empfangen wurden
    Do
        DoEvents
        GetReply Reply
    Loop While Reply.MainState = repPreliminary
    If Reply.MainState <> repComplete Then
        Receive = False
        Exit Function
    End If
    
    'Datei speichern
    If TypeName(LocalFile) = "InetBuffer" Then
        'Empfangende Daten in den Speicher kopieren
        m_Buffer.GetData tmp
        Set LocalBuffer = LocalFile
        LocalBuffer.Clear
        LocalBuffer.Append tmp
    Else
        'Daten in einer lokalen Datei speichern
        tmp = LocalFile
        m_Buffer.SaveFile tmp
    End If

    'Empfangs Buffer löschen
    m_Buffer.Clear
End Function

'Initialisiert ein Datentransfer vom Client zum Server
'Server <- Client
Public Function Store(LocalFile As Variant, ServerFile As String) As Boolean
    'Variablen deklaration
    Dim Buffer As String
    Dim Reply As FtpReply
    Dim tmp As String
    
    'Eingabeparameter überprüfen
    If Not (TypeName(LocalFile) = "String" Or TypeName(LocalFile) = "InetBuffer") Then
        Err.Raise 13  'Typen unverträglich
    End If
    
    'Datei übertragen
    DataInit
    m_Buffer.Clear
    Set Reply = DoExecute("STOR " & ServerFile)
    
    If Reply.MainState <> repPreliminary Then
        'Speichern hat ein Fehler verursacht
        Store = False
        Exit Function
    End If
    
    'Datei öffnen
    Dim ff As Integer
    ff = FreeFile
    Open LocalFile For Binary As #ff
    
    'Solange lesen bis Dateiende erreicht ist
    Do While LOF(ff) > Loc(ff)
        If LOF(ff) > Loc(ff) Then
            'Dateiende noch nicht erreicht, weiterlesen
            If LOF(ff) - Loc(ff) < 8192 Then
                'Zulesende Daten < 8192
                Buffer = Space(LOF(ff) - Loc(ff))
            Else
                'Mehr als 8192 Daten in der Datei
                Buffer = Space(8192)
            End If
            'Daten lesen
            Get #ff, , Buffer
            'Daten senden
            If Not DataSend(Buffer) Then
                Store = False
                Close #ff
                Exit Function
            End If
        End If
    Loop
    'Datei schließen
    Close #ff
    
    'Daten Verbindung trennen
    DataClose
    
    'Reply abwarten
    Do
        DoEvents
        GetReply Reply
    Loop While Reply.MainState = repPreliminary
    Store = (Reply.MainState = repComplete)
End Function
