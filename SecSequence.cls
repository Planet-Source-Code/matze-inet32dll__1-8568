VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Benutzerkennung
Private m_Username As String
Private m_Password As String
Private m_Domain As String

'Package Information
Private m_Package As New SecPackage
Private m_IsInSequence As Boolean

'Token Ausgabe
Private m_Token As String

'Security Handles
Private m_hCredentials As DWORD
Private m_hContext As DWORD


'GET Methoden
Public Property Get Package() As Variant: Package = m_Package.Package: End Property
Public Property Get Username() As String: Username = m_Username: End Property
Public Property Get Password() As String: Password = m_Password: End Property
Public Property Get Domain() As String: Domain = m_Domain: End Property

'Benutzernamen setzen
Public Property Let Username(NewUsername As String)
    'Nur Schreibzugriff erlauben, wenn nicht in der Autorisierungssequenz
    If Not m_IsInSequence Then
        m_Username = NewUsername
    Else
        Err.Raise vbObjectError, "Inet32.SecSequence", "Diese Eigenschaft ist im aktuellen Zustand des Objektes schreibgeschützt"
    End If
End Property

'Kennwort setzen
Public Property Let Password(NewPassword As String)
    'Nur Schreibzugriff erlauben, wenn nicht in der Autorisierungssequenz
    If Not m_IsInSequence Then
        m_Password = NewPassword
    Else
        Err.Raise vbObjectError, "Inet32.SecSequence", "Diese Eigenschaft ist im aktuellen Zustand des Objektes schreibgeschützt"
    End If
End Property

'Domäne setzen
Public Property Let Domain(NewDomain As String)
    'Nur Schreibzugriff erlauben, wenn nicht in der Autorisierungssequenz
    If Not m_IsInSequence Then
        m_Domain = NewDomain
    Else
        Err.Raise vbObjectError, "Inet32.SecSequence", "Diese Eigenschaft ist im aktuellen Zustand des Objektes schreibgeschützt"
    End If
End Property

'Ändert das benutzte Sicherheitspacket für die Autorisierung
Public Property Let Package(NewPackage As Variant)
    'Nur Schreibzugriff erlauben, wenn nicht in der Autorisierungssequenz
    If Not m_IsInSequence Then
        If TypeName(NewPackage) = "SecPackage" Then
            Set Package = NewPackage
        Else
            Set m_Package = New SecPackage
            m_Package.Package = NewPackage
        End If
    Else
        Err.Raise vbObjectError, "Inet32.SecSequence", "Diese Eigenschaft ist in der ""Authorization-Sequence"" schreib geschützt."
    End If
End Property


'Ermittelt ob der Vorgang der Autorisierung noch in Gang ist
Public Function IsInSequence() As Boolean
    IsInSequence = m_IsInSequence
End Function


'Initialisiet oder führt eine Sequenz fort
Public Function DoSequence(Optional InToken As String, Optional OutToken As String) As Boolean
    'Variablen deklaration
    Dim AuthIdentity As SEC_WINNT_AUTH_IDENTITY
    Dim UsernameBuf(20) As Byte
    Dim PasswordBuf(20) As Byte
    Dim DomainBuf(20) As Byte
    Dim OutBuffDesc As SecBufferDesc
    Dim OutSecBuff As SecBuffer
    Dim InBuffDesc As SecBufferDesc
    Dim InSecBuff As SecBuffer
    Dim ContextAttributes As Long
    Dim LifeTime As DWORD
    Dim nRet As Long

    'Prüfen on die Benutzer Eigenschaft gesetzt ist
    If m_Username = "" Then
        Err.Raise vbObjectError, "Inet32.SecSequence", "Bevor eine Sequence initialisiert werden kann. Muss die Eigenschaft auf Username auf einen Wert gesetzt werden."
    End If

    'Prüfen welches Package verwendet werden soll
    If UCase(m_Package.Package) = "BASIC" Then
        'Basic Package benutzen
        If m_Domain = "" Then
            m_Token = m_Username & ":" & m_Password
        Else
            m_Token = m_Domain & "\" & m_Username & ":" & m_Password
        End If
        OutToken = m_Token
        m_IsInSequence = False
        DoSequence = True
    Else
        'Security Service Provider Interface benutzen
        If InToken = "" And m_IsInSequence Then
            'Nicht zulässig in der Authorisierungs Sequenz kein InToken
            Err.Raise secInvalidData, "HttpRequest.HttpSecurity", "Keine Daten in der Autorisierungssequenz"
        Else
            If Not m_IsInSequence Then
                'Sequenze initialisieren
                If m_Domain = "" Then
                    Err.Raise vbObjectError, "Inet32.SecSequence", "Die Domain Eigenschaft ist ungültig"
                End If
            
                'Benutzer Autorisierungs Struktur
                AuthIdentity.User = VarPtr(UsernameBuf(0))
                AuthIdentity.UserLength = Len(m_Username)
                AuthIdentity.Password = VarPtr(PasswordBuf(0))
                AuthIdentity.PasswordLength = Len(m_Password)
                AuthIdentity.Domain = VarPtr(DomainBuf(0))
                AuthIdentity.DomainLength = Len(m_Domain)
                AuthIdentity.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI
            
                'Strings in Byte Arrays konvertieren
                StrToByte m_Username, UsernameBuf
                StrToByte m_Password, PasswordBuf
                StrToByte m_Domain, DomainBuf
                
                'Output Buffer vorbereiten
                ReDim pOut(m_Package.MaxToken) As Byte
                OutSecBuff.cbBuffer = m_Package.MaxToken
                OutSecBuff.pvBuffer = VarPtr(pOut(0))
                OutSecBuff.BufferType = SECBUFFER_TOKEN
                OutBuffDesc.ulVersion = 0
                OutBuffDesc.cBuffers = 1
                OutBuffDesc.pBuffers = VarPtr(OutSecBuff)
                
                'Benutzer Konto anfordern
                nRet = AcquireCredentialsHandle(0, m_Package.Package, SECPKG_CRED_OUTBOUND, 0, _
                       VarPtr(AuthIdentity), 0, 0, m_hCredentials, LifeTime)
                If nRet < 0 Then GoTo Error
                
                'Token initialisieren
                nRet = InitializeSecurityContext(m_hCredentials, 0, "Inet32.dll", 0, 0, _
                       SECURITY_NATIVE_DREP, 0, 0, m_hContext, OutBuffDesc, _
                       ContextAttributes, LifeTime)
                If nRet < 0 Then GoTo Error
        
                'Token komplettieren
                If nRet = SEC_I_COMPLETE_NEEDED Or nRet = SEC_I_COMPLETE_AND_CONTINUE Then
                    nRet = CompleteAuthToken(m_hContext, OutBuffDesc)
                    If nRet < 0 Then GoTo Error
                End If
        
                'Byte Array in ein String konvertieren
                ByteToStr pOut, m_Token, OutSecBuff.cbBuffer
                
                'Output Token zurück geben
                m_IsInSequence = True
                OutToken = m_Token
                DoSequence = True
                
            Else
                'Autorisierungs Sequenze fortsetzen
                
                'Array Größen an maximale Größe eines Tokens anpassen
                ReDim pIn(m_Package.MaxToken) As Byte
                ReDim pOut(m_Package.MaxToken) As Byte
                'Input Buffer vorbereiten
                InSecBuff.cbBuffer = m_Package.MaxToken
                InSecBuff.pvBuffer = VarPtr(pIn(0))
                InSecBuff.BufferType = SECBUFFER_TOKEN
                InBuffDesc.ulVersion = 0
                InBuffDesc.cBuffers = 1
                InBuffDesc.pBuffers = VarPtr(InSecBuff)
                'Output Buffer vorbereiten
                OutSecBuff.cbBuffer = m_Package.MaxToken
                OutSecBuff.pvBuffer = VarPtr(pOut(0))
                OutSecBuff.BufferType = SECBUFFER_TOKEN
                OutBuffDesc.ulVersion = 0
                OutBuffDesc.cBuffers = 1
                OutBuffDesc.pBuffers = VarPtr(OutSecBuff)
               
                'Input Token decodieren
                InSecBuff.cbBuffer = Len(InToken)
                StrToByte InToken, pIn
               
                'Token verarbeiten
                nRet = InitializeSecurityContext(m_hCredentials, VarPtr(m_hContext), "Inet32.dll", 0, 0, _
                       SECURITY_NATIVE_DREP, VarPtr(InBuffDesc), 0, m_hContext, OutBuffDesc, _
                       ContextAttributes, LifeTime)
                If nRet < 0 Then GoTo Error
        
                'Token komplettieren
                If nRet = SEC_I_COMPLETE_NEEDED Or nRet = SEC_I_COMPLETE_AND_CONTINUE Then
                    nRet = CompleteAuthToken(m_hContext, OutBuffDesc)
                    If nRet < 0 Then GoTo Error
                End If
                
                'Byte Array in ein String konvertieren
                ByteToStr pOut, m_Token, OutSecBuff.cbBuffer
                
                'Output Token zurückgeben
                m_IsInSequence = (nRet <> SEC_E_OK)
                OutToken = m_Token
                DoSequence = True
            End If
        End If
    End If
    Exit Function
    
Error:
    'Fehlerbehandeln
    Err.Raise vbObjectError, "Inet32.SecSequence", GetMsg(nRet)
    m_Token = ""
    OutToken = ""
    DoSequence = False
End Function


'Token Output, dises Token muss zum Server geschickt werden
Public Property Get Token() As String
    Token = m_Token
End Property
