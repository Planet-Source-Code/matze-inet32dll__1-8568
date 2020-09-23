VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Data Members
Private m_Resource As New InetURL
Private m_SecSequence As New SecSequence
Private m_UseForProxy As Boolean
Private m_Header As HttpHeader
Private m_Token As String
Private m_IsInSequence As Boolean

'Resource Name zurückliefern für die, die Autorisierung gemacht wird
Public Property Get Resource() As InetURL: Set Resource = m_Resource: End Property

'Resource Name setzen für die die Autorisierung gemacht werden soll
Public Property Let Resource(NewResource As Variant)
    'Eigenschaft ist während des Autorisierungs Vorgangs schreibgeschützt
    If Not m_IsInSequence Then
        'Eingabe Parameter prüfen
        If TypeName(NewResource) = "InetURL" Then
            Set m_Resource = NewResource
        ElseIf TypeName(NewResource) = "String" Then
            Set m_Resource = New InetURL
            m_Resource.Url = NewResource
        Else
            Err.Raise vbObjectError, "Inet32.HttpSecurity", "Unerwarteter Parametertyp. Nur InetURL und String sind erlaubt"
        End If
        
        'Werte auf Security Objekt übertragen
        m_SecSequence.Username = m_Resource.Username
        m_SecSequence.Password = m_Resource.Password
        m_SecSequence.Domain = m_Resource.Domain
    Else
        Err.Raise vbObjectError, "Diese Operation ist im aktuellen Zustands des Objektes nicht erlaubt."
    End If
End Property


'Bestimmt ob die Autorisierung für einen Proxy oder Www Site gemacht werden
Public Property Get UseForProxy() As String: UseForProxy = m_UseForProxy: End Property
Public Property Let UseForProxy(NewUseForProxy As String): m_UseForProxy = NewUseForProxy: End Property


'Gibt zurück ob wir uns augenblicklich in der Autorisierungssequence befinden
Public Function IsInSequence() As Boolean
    IsInSequence = m_IsInSequence
End Function


'Input Token vom Server System einlesen
Public Function HttpAuthenticate(Header As HttpHeader) As Boolean
    'Variablen deklaration
    Dim Token As String
    Dim IsOk As Boolean
    Dim i As Integer
    
    If Not m_IsInSequence Then
        'Sequence initialisieren
        Set m_Header = Header
        HttpAuthenticate = True
        
    Else
        'Sequence fortsetzen
        Set m_Header = Header
        
        'Prüfen ob die Package Namen übereinstimmen
        IsOk = False
        If m_UseForProxy Then
            Do While m_Header.ProxyAuthenticate(i).Package <> ""
                If UCase(m_Header.ProxyAuthenticate(i).Package) = UCase(m_SecSequence.Package) Then
                    Token = m_Header.ProxyAuthenticate(i).Data
                    IsOk = True
                End If
                i = i + 1
            Loop
        Else
            Do While m_Header.Authenticate(i).Package <> ""
                If UCase(m_Header.Authenticate(i).Package) = UCase(m_SecSequence.Package) Then
                    Token = m_Header.Authenticate(i).Data
                    IsOk = True
                End If
                i = i + 1
            Loop
        End If
             
        If IsOk Then
            'Sequence fortsetzen
            If UCase(m_SecSequence.Package) = "BASIC" Then
                HttpAuthenticate = m_SecSequence.DoSequence(, m_Token)
            Else
                HttpAuthenticate = m_SecSequence.DoSequence(Base64Decode(Token), m_Token)
            End If
        Else
            Err.Raise vbObjectError, "Inet.HttpSecurity", "Die Package Namen stimmen nicht überein"
        End If
    End If
End Function


Public Function HttpAuthorization(AuthorizationHeader As String) As Boolean
    'Variablen deklaration
    Dim Packages As SecPackages
    Dim Token As String
    Dim NoMatch As Boolean
    Dim i As Integer
    Dim n As Integer
    
    If Not m_IsInSequence Then
        'Lokal installierte Package Liste erstellen
        Set Packages = New SecPackages
        
        'Übereinstimmung suchen, mit der über den Server kommuniziert werden kann
        NoMatch = True
        If m_UseForProxy Then
            Do While m_Header.ProxyAuthenticate(i).Package <> ""
                For n = 0 To Packages.Count - 1
                    If UCase(m_Header.ProxyAuthenticate(i).Package) = UCase(Packages(n).Package) Then
                        NoMatch = False
                        m_SecSequence.Package = Packages(n).Package
                        Exit For
                    End If
                Next n
                i = i + 1
            Loop
        Else
            Do While m_Header.Authenticate(i).Package <> ""
                For n = 0 To Packages.Count - 1
                    If UCase(m_Header.Authenticate(i).Package) = UCase(Packages(n).Package) Then
                        NoMatch = False
                        m_SecSequence.Package = Packages(n).Package
                        Exit For
                    End If
                Next n
                i = i + 1
            Loop
        End If
        
        'Prüfen, ob eine übereinstimmung gefunden wurde
        If Not NoMatch Then
            'Sequence initialisieren und
            'HTTP Header bilden
            If m_SecSequence.DoSequence(, Token) Then
                If m_UseForProxy Then
                    AuthorizationHeader = AuthorizationHeader & "Proxy-Authorization: " & m_SecSequence.Package & " " & Base64Encode(Token) & vbCrLf
                Else
                    AuthorizationHeader = AuthorizationHeader & "Authorization: " & m_SecSequence.Package & " " & Base64Encode(Token) & vbCrLf
                End If
                m_IsInSequence = m_SecSequence.IsInSequence
                HttpAuthorization = True
            Else
                Err.Raise vbObjectError, "Inet32.HttpSecurity", "Es konnte kein HTTP Authorization Header gebildet werden"
            End If
        Else
            Err.Raise vbObjectError, "Inet32.HttpSecurity", "Keine übereinstimmenden Packages gefunden"
        End If
        
    Else
        'Token zurueckgeben
        If m_UseForProxy Then
            AuthorizationHeader = AuthorizationHeader & "Proxy-Authorization: " & m_SecSequence.Package & " " & Base64Encode(m_Token) & vbCrLf
        Else
            AuthorizationHeader = AuthorizationHeader & "Authorization: " & m_SecSequence.Package & " " & Base64Encode(m_Token) & vbCrLf
        End If
        m_IsInSequence = m_SecSequence.IsInSequence
        HttpAuthorization = True
    End If
End Function


'Konstruktion
Private Sub Class_Initialize()
    'Variablen initialisieren
    m_SecSequence.Package = "NTLM"
End Sub
