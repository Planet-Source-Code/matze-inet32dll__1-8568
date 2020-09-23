VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Object References
Private WithEvents tmrReading As Timer
Attribute tmrReading.VB_VarHelpID = -1
Private frxFileContainer As New frmFileContainer

'Data Members
Private m_FileNumber As Integer
Private m_Header As New FileHeader
Private m_Buffer As New InetBuffer
Private m_Blocking As Boolean
Private m_State As Integer
Private m_Tag As Variant
Private m_Resource As New InetURL

'Events
Event Request(ByVal Resource As InetURL)
Event Progress(ByVal bytesReceived As Long, ByVal bytesTotal As Long)
Event Complete()

Public Property Get Header() As FileHeader: Set Header = m_Header: End Property
Public Property Get Buffer() As InetBuffer: Set Buffer = m_Buffer: End Property
Public Property Get Blocking() As Boolean: Blocking = m_Blocking: End Property
Public Property Get Resource() As InetURL: Set Resource = m_Resource: End Property
Public Property Get State() As InetStates: State = m_State: End Property

'Tag Property Implementieren
Public Property Get Tag() As Variant: Set Tag = m_Tag: End Property
Public Property Let Tag(NewTag As Variant): m_Tag = NewTag: End Property
Public Property Set Tag(NewTag As Variant): Set m_Tag = NewTag: End Property


Public Function OpenURL(Resource As Variant, Optional Sync As Boolean) As InetBuffer
    'Variablen deklaration
    Dim tmpResource As InetURL
    
    'Prüfen ob Objekt bereit für Anforderung ist
    If m_Blocking Then
        Err.Raise vbObjectError
    End If
    m_Blocking = True
    
    'Eingabe Parameter prüfen
    If TypeName(Resource) = "InetURL" Then
        Set tmpResource = Resource
    ElseIf TypeName(Resource) = "String" Then
        Set tmpResource = New InetURL
        tmpResource = Resource
    Else
        Err.Raise vbObjectError
    End If
    
    'Prüfen ob es sich um einen 'file' Link handelt
    If LCase(tmpResource.Protocol) <> "file" Then
        Err.Raise vbObjectError
    End If
        
    'Data Members neu initialisieren
    Set m_Header = New FileHeader
    Set m_Buffer = New InetBuffer
    
    'Datei öffnen
    On Error Resume Next
    m_FileNumber = FreeFile
    Open tmpResource.UrlPath For Binary Access Read As #m_FileNumber
    
    'Prüfen ob ein Fehler aufgetreten ist
    If Err.Number <> 0 Then
        'Datei wurde nicht gefunden
        m_Blocking = False
        m_Header.Status = inetNotFound
        RaiseEvent Complete
    Else
        'Timer initialisieren
        tmrReading.Interval = 500
        tmrReading.Enabled = True
        'Event auslösen
        RaiseEvent Request(tmpResource)
    End If
    
    'Header Informationen ausfüllen
    m_Header.Attributes = GetAttr(tmpResource.UrlPath)
    m_Header.Length = LOF(m_FileNumber)
    m_Header.Request = Now
    m_Header.Modified = Now
    
    If Sync Then
        'Syncron warten bis Anforderung beendet
        Do While m_Blocking
            DoEvents
        Loop
        'Buffer Objekt zurückgeben
        Set OpenURL = m_Buffer
    Else
        'Leeres Objekt zurückgeben
        Set OpenURL = Nothing
    End If
End Function

'Lese Vorgang abbrechen
Public Sub Cancel()
    'Anforderung abbrechen
    If m_Blocking Then
        Close #m_FileNumber
        m_Blocking = False
        tmrReading.Enabled = False
        m_Header.Status = inetCanceled
        RaiseEvent Complete
    End If
End Sub

'Initialisieren
Private Sub Class_Initialize()
    Set tmrReading = frmFileContainer.tmrFile
End Sub

'Datei im Hintergrund lesen
Private Sub tmrReading_Timer()
    'Daten auslesen
    Dim Buffer As String
    
    If LOF(m_FileNumber) > Loc(m_FileNumber) Then
        'Dateiende noch nicht erreicht, weiterlesen
        If LOF(m_FileNumber) - Loc(m_FileNumber) < 8192 Then
            'Zulesende Daten < 8192
            Buffer = Space(LOF(m_FileNumber) - Loc(m_FileNumber))
        Else
            'Mehr als 8192 Daten in der Datei
            Buffer = Space(8192)
        End If
        'Daten lesen
        Get #m_FileNumber, , Buffer
        'Buffer füllen
        m_Buffer.Append Buffer
        RaiseEvent Progress(m_Buffer.Length, LOF(m_FileNumber))
    End If
    
    'Prüfen ob das Dateiende erreicht ist
    If LOF(m_FileNumber) <= Loc(m_FileNumber) Then
        'Dateiende erreicht
        m_Blocking = False
        tmrReading.Enabled = False
        Close #m_FileNumber
        m_Header.Status = inetOK
        RaiseEvent Complete
    End If
End Sub
