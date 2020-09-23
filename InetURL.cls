VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InetURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"URL"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Protocol As String
Private m_Username As String
Private m_Password As String
Private m_Hostname As String
Private m_Hostport As Integer
Private m_UrlPath As String


Public Property Let Url(Resource As String)
    'Protokoll aus der Url parsen
    If lPos(Resource, "://") = 0 Then
        m_Protocol = "http"
    Else
        m_Protocol = LCase(Left(Resource, lPos(Resource, "://") - 1))
        Resource = Mid(Resource, Len(m_Protocol) + 4)
    End If
    
    'Parser für Protokoll wählen
    Select Case m_Protocol
        Case "http", "ftp"
            'Benutzername und Kennwort aus der Url aus Parsen (falls vorhanden)
            Dim tmp As String
            If lPos(Resource, "/") <> 0 Then
                tmp = Left(Resource, lPos(Resource, "/"))
                Resource = Mid(Resource, lPos(Resource, "/") + 1)
            Else
                tmp = Resource
                Resource = ""
            End If
            If lPos(tmp, "@") = 0 Then
                m_Username = ""
                m_Password = ""
            Else
                If lPos(tmp, ":") > 0 And lPos(tmp, ":") < lPos(tmp, "@") Then
                    m_Username = Left(tmp, lPos(tmp, ":") - 1)
                    tmp = Mid(tmp, lPos(tmp, ":") + 1)
                    m_Password = Left(tmp, lPos(tmp, "@") - 1)
                    tmp = Mid(tmp, lPos(tmp, "@") + 1)
                Else
                    m_Username = Left(tmp, lPos(tmp, "@") - 1)
                    tmp = Mid(tmp, lPos(tmp, "@") + 1)
                    m_Password = ""
                End If
            End If
            Resource = tmp & Resource
        
            'UrlPath aus der Url parsen
            If lPos(Resource, "/") <> 0 Then
                Me.UrlPath = Mid(Resource, lPos(Resource, "/"))
                Resource = Left(Resource, lPos(Resource, "/") - 1)
            Else
                m_UrlPath = "/"
            End If
            
            'Server und Port aus der Url parsen
            If lPos(Resource, ":") = 0 Then
                m_Hostport = 0
                If Protocol = "http" Then m_Hostport = 80
                If Protocol = "ftp" Then m_Hostport = 21
                m_Hostname = Resource
            Else
                m_Hostname = Left(Resource, lPos(Resource, ":") - 1)
                m_Hostport = Mid(Resource, lPos(Resource, ":") + 1)
            End If
    
        Case "file"
            m_UrlPath = Resource
        
        Case Else
            Err.Raise vbObjectError, "Inet32.InetURL", "Nicht unterstütztes Protokoll kann String nicht auflösen"
    End Select
End Property

Public Property Get Url() As String
Attribute Url.VB_UserMemId = 0
    'Variablen deklaration
    Dim tmp As String
    
    'Die Url dem Protocol entsprechend wieder bilden
    Select Case m_Protocol
        Case ""
            'Ein Leerstring zurückgeben, da Objekt nicht richtig initialisiert
            Url = ""
            
        Case "http", "ftp"
            If m_Hostname = "" Then
                'Kein Server angegeben, relative Url bilden
                tmp = m_UrlPath
            Else
                'Absolute Url bilden
                tmp = m_Protocol + "://"
                If m_Username <> "" And Right(m_Username, 1) <> "\" Then
                    tmp = tmp & m_Username
                    If m_Password <> "" Then tmp = tmp & ":" & m_Password
                    tmp = tmp & "@"
                End If
                tmp = tmp + m_Hostname
                If m_Hostport <> 80 Then tmp = tmp & ":" & m_Hostport
                tmp = tmp & m_UrlPath
            End If
            Url = tmp
            
        Case "file"
            tmp = m_Protocol + "://" & m_UrlPath
            Url = tmp
            
        Case Else
            Err.Raise vbObjectError, "Inet32.InetURL", "Nicht unterstütztes Protokoll"
    End Select
End Property
Public Property Get Protocol() As String: Protocol = m_Protocol: End Property
Public Property Get Password() As String: Password = m_Password: End Property
Public Property Get Hostname() As String: Hostname = m_Hostname: End Property
Public Property Get Hostport() As Integer: Hostport = m_Hostport: End Property
Public Property Get UrlPath() As String: UrlPath = m_UrlPath: End Property


Private Function RemoveDots(DotString As String) As String
    Dim NewPath As String
    Dim tmpPath As String
    
    NewPath = DotString
    'Die "." Verzeichnisse aus dem Dokumenten Pfad entfernen
    Do While lPos(NewPath, "/./") <> 0
        NewPath = Left(NewPath, lPos(NewPath, "/./")) & Mid(NewPath, lPos(NewPath, "/./") + 3)
    Loop
        
    'Die ".." Verzeichnisse aus dem Dokumenten Pfad auflösen
    Do While lPos(NewPath, "/../") <> 0
        tmpPath = Left(NewPath, lPos(NewPath, "/../") - 1)
        tmpPath = Left(tmpPath, rPos(tmpPath, "/"))
        NewPath = tmpPath & Mid(NewPath, lPos(NewPath, "/../") + 4)
    Loop
    
    'String zurückgeben
    RemoveDots = NewPath
End Function

'Verändert den UrlPath
Public Property Let UrlPath(ByVal NewUrlPath As String)
    'Variablen deklaration
    Dim NewPath As String
    Dim tmpPath As String
    Dim Query As String
    Dim Textmark As String
    
    NewPath = NewUrlPath
    'Query und/oder Textmarke vom Dokumenten Namen trennen
    If lPos(NewPath, "?") <> 0 Then
        Query = Mid(NewPath, lPos(NewPath, "?"))
        NewPath = Left(NewPath, lPos(NewPath, "?") - 1)
    End If
    If lPos(NewPath, "#") <> 0 Then
        Textmark = Mid(NewPath, lPos(NewPath, "#"))
        NewPath = Left(NewPath, lPos(NewPath, "#") - 1)
    End If

    'Überprüfen ob es sich nicht doch um eine absolute Url handelt
    If lPos(NewPath, "://") <> 0 Then
        'Eine absolute Url, auflösen durch Url Eigenschaft
        Me.Url = NewPath & Textmark & Query
    Else
        'Nur der UrlPath, versuchen den UrlPath aufzulösen
        'Prüfen ob es sich um einen absoluten UrlPath handelt
        If Left(NewPath, 1) = "/" Then
            'Ein absoluter UrlPath
            m_UrlPath = RemoveDots(NewPath) & Textmark & Query
        Else
            'Relativer UrlPath
            m_UrlPath = RemoveDots(Me.Path & NewPath) & Textmark & Query
        End If
    End If
End Property

'Liefert den Pfad aus dem UrlPath
Public Property Get Path() As String
    'Query und/oder Textmarke vom Dokumenten Namen trennen
    Dim RetPath As String
    RetPath = m_UrlPath
    If lPos(RetPath, "?") <> 0 Then RetPath = Left(RetPath, lPos(RetPath, "?") - 1)
    If lPos(RetPath, "#") <> 0 Then RetPath = Left(RetPath, lPos(RetPath, "#") - 1)
    If Right(RetPath, 1) <> "/" Then RetPath = Left(RetPath, rPos(RetPath, "/"))
    Path = RetPath
End Property

'Liefert das Dokument aus dem UrlPath
Public Property Get Document() As String
    'Query und/oder Textmarke vom Dokumenten Namen trennen
    Dim RetDocument As String
    RetDocument = m_UrlPath
    If lPos(RetDocument, "?") <> 0 Then RetDocument = Left(RetDocument, lPos(RetDocument, "?") - 1)
    If lPos(RetDocument, "#") <> 0 Then RetDocument = Left(RetDocument, lPos(RetDocument, "#") - 1)
    RetDocument = Mid(RetDocument, rPos(RetDocument, "/") + 1)
    Document = RetDocument
End Property

'Liefert die Query aus dem UrlPath
Public Property Get Query() As String
    If lPos(m_UrlPath, "?") <> 0 Then
        Query = Mid(m_UrlPath, lPos(m_UrlPath, "?"))
    Else
        Query = ""
    End If
End Property

'Liefert die Textmarke aus dem UrlPath
Public Property Get Textmark() As String
    If lPos(m_UrlPath, "#") <> 0 Then
        Textmark = Mid(m_UrlPath, lPos(m_UrlPath, "#"))
    Else
        Textmark = ""
    End If
End Property

'Verändert die Textmarke im UrlPath
Public Property Let Textmark(NewTextmark As String)
    Dim tmp As String
    If NewTextmark = "" Then
        m_UrlPath = Me.Path & Me.Document & Me.Query
    Else
        tmp = NewTextmark
        If Left(tmp, 1) <> "#" Then tmp = "#" & tmp
        m_UrlPath = Me.Path & Me.Document & tmp & Me.Query
    End If
End Property

'Benutzernamen aus der Url parsen
Public Property Get Username() As String
    If lPos(m_Username, "\") <> 0 Then
        Username = Mid(m_Username, lPos(m_Username, "\") + 1)
    Else
        Username = m_Username
    End If
End Property

'Liefert die Domain zurück
Public Property Get Domain() As String
    If lPos(m_Username, "\") <> 0 Then
        Domain = Left(m_Username, lPos(m_Username, "\") - 1)
    Else
        Domain = ""
    End If
End Property

'Setzt einen neuen Benutzernamen
Public Property Let Username(NewUsername As String)
    'Variablen deklaration
    Dim tmp As String
    
    'Domäne abtrennen, falls ein Benutzername in der Form DOM\USER übergeben wurde
    If lPos(NewUsername, "\") <> 0 Then
        tmp = Mid(NewUsername, lPos(NewUsername, "\") + 1)
    Else
        tmp = NewUsername
    End If
    
    'Benutzername speichern
    If Domain = "" Then
        m_Username = NewUsername
    Else
        m_Username = Domain & "\" & m_Username
    End If
End Property

'Setzt ein neuen Domain Namen
Public Property Let Domain(NewDomain As String)
    'Variablen deklaration
    Dim tmp As String
    
    'Benutzername abtrennen, falls ein Benutzername in der Form DOM\USER übergeben wurde
    If lPos(NewDomain, "\") <> 0 Then
        tmp = Left(NewDomain, lPos(NewDomain, "\") - 1)
    Else
        tmp = NewDomain
    End If
    
    'Domäne speichern
    m_Username = tmp & "\" & Username
End Property

'Setzt ein neues Passwort
Public Property Let Password(NewPassword As String)
    m_Password = NewPassword
End Property
