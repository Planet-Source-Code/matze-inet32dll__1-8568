Attribute VB_Name = "basSSPI"
'**************************************
' Windows API/Global Declarations for
' Validate User Credentials on 95/98/NT with NTLM
'**************************************

Option Explicit
Public Const SEC_E_OK = &H0
Public Const SEC_E_INSUFFICIENT_MEMORY = &H80090300
Public Const SEC_E_INVALID_HANDLE = &H80090301
Public Const SEC_E_UNSUPPORTED_FUNCTION = &H80090302
Public Const SEC_E_TARGET_UNKNOWN = &H80090303
Public Const SEC_E_INTERNAL_ERROR = &H80090304
Public Const SEC_E_SECPKG_NOT_FOUND = &H80090305
Public Const SEC_E_NOT_OWNER = &H80090306
Public Const SEC_E_CANNOT_INSTALL = &H80090307
Public Const SEC_E_INVALID_TOKEN = &H80090308
Public Const SEC_E_CANNOT_PACK = &H80090309
Public Const SEC_E_QOP_NOT_SUPPORTED = &H8009030A
Public Const SEC_E_NO_IMPERSONATION = &H8009030B
Public Const SEC_E_LOGON_DENIED = &H8009030C
Public Const SEC_E_UNKNOWN_CREDENTIALS = &H8009030D
Public Const SEC_E_NO_CREDENTIALS = &H8009030E
Public Const SEC_E_MESSAGE_ALTERED = &H8009030F
Public Const SEC_E_OUT_OF_SEQUENCE = &H80090310
Public Const SEC_E_NO_AUTHENTICATING_AUTHORITY = &H80090311
Public Const SEC_I_CONTINUE_NEEDED = &H90312
Public Const SEC_I_COMPLETE_NEEDED = &H90313
Public Const SEC_I_COMPLETE_AND_CONTINUE = &H90314
Public Const SEC_I_LOCAL_LOGON = &H90315
Public Const SEC_E_BAD_PKGID = &H80090316
Public Const SEC_E_CONTEXT_EXPIRED = &H80090317
Public Const SEC_E_INCOMPLETE_MESSAGE = &H80090318
Public Const SEC_E_INCOMPLETE_CREDENTIALS = &H80090320
Public Const SEC_E_BUFFER_TOO_SMALL = &H80090321
Public Const SEC_I_INCOMPLETE_CREDENTIALS = &H90320
Public Const SEC_I_RENEGOTIATE = &H90321
Public Const SEC_E_WRONG_PRINCIPAL = &H80090322
Public Const SECPKG_CRED_OUTBOUND = 2
Public Const SECPKG_CRED_INBOUND = 1
Public Const SEC_WINNT_AUTH_IDENTITY_ANSI = 1
Public Const SEC_WINNT_AUTH_IDENTITY_UNICODE = 2
Public Const SECURITY_NATIVE_DREP = 16
Public Const SECURITY_NETWORK_DREP = 0
Public Const SECBUFFER_TOKEN = 2


Public Type SecPkgInfo
    fCapabilities As Long 'unsigned long Capability bitmask
    wVersion As Integer 'unsigned short Version of driver
    wRPCID As Integer 'unsigned short ID For RPC Runtime
    cbMaxToken As Long 'unsigned long Size of authentication token (max)
    Name As Long 'SEC_CHAR SEC_FAR * Text name
    Comment As Long 'SEC_CHAR SEC_FAR * Comment
End Type


Public Type SEC_WINNT_AUTH_IDENTITY
    User As Long 'unsigned char __RPC_FAR *
    UserLength As Long 'unsigned long
    Domain As Long 'unsigned char __RPC_FAR *
    DomainLength As Long 'unsigned long
    Password As Long 'unsigned char __RPC_FAR *
    PasswordLength As Long 'unsigned long
    Flags As Long 'unsigned long
End Type


Public Type DWORD
    dwLower As Long 'unsigned long
    dwUpper As Long 'unsigned long
End Type


Public Type SecBuffer
    cbBuffer As Long 'unsigned long Size of the buffer, in bytes
    BufferType As Long 'unsigned long Type of the buffer (below)
    pvBuffer As Long 'void SEC_FAR * Pointer to the buffer
End Type


Public Type SecBufferDesc
    ulVersion As Long 'unsigned long Version number
    cBuffers As Long 'unsigned long Number of buffers
    pBuffers As Long 'PSecBuffer Pointer to array of buffers
End Type


Private Declare Function AcquireCredentialsHandleNT Lib "security.dll" _
    Alias "AcquireCredentialsHandleA" ( _
    ByVal pszPrincipal As Long, ByVal pszPackage As String, _
    ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, _
    ByVal pvGetKeyArgument As Long, ByRef PCredHandle As DWORD, _
    ByRef ptsExpiry As DWORD) As Long
Private Declare Function AcquireCredentialsHandle9X Lib "secur32.dll" _
    Alias "AcquireCredentialsHandleA" ( _
    ByVal pszPrincipal As Long, ByVal pszPackage As String, _
    ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, _
    ByVal pvGetKeyArgument As Long, ByRef PCredHandle As DWORD, _
    ByRef ptsExpiry As DWORD) As Long


Private Declare Function InitializeSecurityContextNT Lib "security.dll" _
    Alias "InitializeSecurityContextA" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByVal pszTargetName As String, ByVal fContextReq As Long, _
    ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, _
    ByRef phNewContext As DWORD, ByRef pOutput As SecBufferDesc, _
    ByRef pfContextAttr As Long, ByRef ptsExpiry As DWORD) As Long
Private Declare Function InitializeSecurityContext9X Lib "secur32.dll" _
    Alias "InitializeSecurityContextA" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByVal pszTargetName As String, ByVal fContextReq As Long, _
    ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, _
    ByRef phNewContext As DWORD, ByRef pOutput As SecBufferDesc, _
    ByRef pfContextAttr As Long, ByRef ptsExpiry As DWORD) As Long


'MyDeclarations
Private Declare Function AcceptSecurityContextNT Lib "security.dll" _
    Alias "AcceptSecurityContext" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
Private Declare Function AcceptSecurityContext9X Lib "secur32.dll" _
    Alias "AcceptSecurityContext" ( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long


Private Declare Function CompleteAuthTokenNT Lib "security.dll" Alias "CompleteAuthToken" _
    (ByRef phContext As DWORD, _
    ByRef pToken As SecBufferDesc) As Long
Private Declare Function CompleteAuthToken9X Lib "secur32.dll" Alias "CompleteAuthToken" _
    (ByRef phContext As DWORD, _
    ByRef pToken As SecBufferDesc) As Long


Private Declare Function DeleteSecurityContextNT Lib "security.dll" Alias "DeleteSecurityContext" _
    (ByRef hCtxt As DWORD) As Long
Private Declare Function DeleteSecurityContext9X Lib "secur32.dll" Alias "DeleteSecurityContext" _
    (ByRef hCtxt As DWORD) As Long


Private Declare Function FreeContextBufferNT Lib "security.dll" Alias "FreeContextBuffer" _
    (ByVal pvContextBuffer As Long) As Long
Private Declare Function FreeContextBuffer9X Lib "secur32.dll" Alias "FreeContextBuffer" _
    (ByVal pvContextBuffer As Long) As Long


Private Declare Function FreeCredentialsHandleNT Lib "security.dll" Alias "FreeCredentialsHandle" _
    (ByRef hCred As DWORD) As Long
Private Declare Function FreeCredentialsHandle9X Lib "secur32.dll" Alias "FreeCredentialsHandle" _
    (ByRef hCred As DWORD) As Long


Private Declare Function InitSecurityInterfaceNT Lib "security.dll" Alias "InitSecurityInterfaceA" _
    () As Long
Private Declare Function InitSecurityInterface9X Lib "secur32.dll" Alias "InitSecurityInterfaceA" _
    () As Long


'PULONG pcPackages,          // receives the number of packages
'PSecPkgInfo *ppPackageInfo  // receives array of information
Private Declare Function EnumerateSecurityPackagesNT Lib "security.dll" Alias "EnumerateSecurityPackagesA" ( _
    ByRef pcPackages As Long, _
    ByRef ppPackageInfo As Long) As Long
Private Declare Function EnumerateSecurityPackages9X Lib "secur32.dll" Alias "EnumerateSecurityPackagesA" ( _
    ByRef pcPackages As Long, _
    ByRef ppPackageInfo As Long) As Long


Private Declare Function QuerySecurityPackageInfoNT Lib "security.dll" Alias "QuerySecurityPackageInfoA" ( _
    ByVal pszPackageName As String, _
    ByRef ppPackageInfo As Long) As Integer
Private Declare Function QuerySecurityPackageInfo9X Lib "secur32.dll" Alias "QuerySecurityPackageInfoA" ( _
    ByVal pszPackageName As String, _
    ByRef ppPackageInfo As Long) As Integer


'**************************************
' Global API/Methods to Use for Win9X and WinNT
'**************************************


'SEC_CHAR SEC_FAR * pszPrincipal,// Nameof principal
'SEC_CHAR SEC_FAR * pszPackage, // Name of package
'unsigned long fCredentialUse,// Flags indicating use
'void SEC_FAR * pvLogonId,// Pointer to logon ID
'void SEC_FAR * pAuthData,// Package specific data
'SEC_GET_KEY_FN pGetKeyFn,// Pointer to GetKey() func
'void SEC_FAR * pvGetKeyArgument,// Value to pass to GetKey()
'PCredHandle phCredential,// (out) CredHandle
'PTimeStamp ptsExpiry// (out) Lifetime (optional)
Public Function AcquireCredentialsHandle(ByVal pszPrincipal As Long, _
    ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonId As Long, _
    ByVal pAuthData As Long, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
    ByRef PCredHandle As DWORD, ByRef ptsExpiry As DWORD) As Long
    If IsNT() Then
        AcquireCredentialsHandle = AcquireCredentialsHandleNT(pszPrincipal, _
        pszPackage, fCredentialUse, pvLogonId, pAuthData, pGetKeyFn, _
        pvGetKeyArgument, PCredHandle, ptsExpiry)
    Else
        AcquireCredentialsHandle = AcquireCredentialsHandle9X(pszPrincipal, _
        pszPackage, fCredentialUse, pvLogonId, pAuthData, pGetKeyFn, _
        pvGetKeyArgument, PCredHandle, ptsExpiry)
    End If
End Function


'PCredHandle phCredential,// Cred to base context
'PCtxtHandle phContext, // Existing context (OPT)
'SEC_CHAR SEC_FAR * pszTargetName,// Name of target
'unsigned long fContextReq, // Context Requirements
'unsigned long Reserved1,// Reserved, MBZ
'unsigned long TargetDataRep,// Data repof target
'PSecBufferDesc pInput, // Input Buffers
'unsigned long Reserved2,// Reserved, MBZ
'PCtxtHandle phNewContext,// (out) New Context handle
'PSecBufferDesc pOutput, // (inout) Output Buffers
'unsigned long SEC_FAR * pfContextAttr, // (out) Context attrs
'PTimeStamp ptsExpiry// (out) Life span (OPT)
Public Function InitializeSecurityContext(ByRef phCredential As DWORD, _
    ByVal phContext As Long, ByVal pszTargetName As String, _
    ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    ByVal pInput As Long, ByVal Reserved2 As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
    If IsNT() Then
        InitializeSecurityContext = InitializeSecurityContextNT(phCredential, _
        phContext, pszTargetName, fContextReq, Reserved1, TargetDataRep, _
        pInput, Reserved2, phNewContext, pOutput, pfContextAttr, ptsExpiry)
    Else
        InitializeSecurityContext = InitializeSecurityContext9X(phCredential, _
        phContext, pszTargetName, fContextReq, Reserved1, TargetDataRep, _
        pInput, Reserved2, phNewContext, pOutput, pfContextAttr, ptsExpiry)
    End If
End Function


'PCredHandle phCredential,// Cred to base context
'PCtxtHandle phContext, // Existing context (OPT)
'PSecBufferDesc pInput, // Input buffer
'unsigned long fContextReq, // Context Requirements
'unsigned long TargetDataRep,// Target Data Rep
'PCtxtHandle phNewContext,// (out) New context handle
'PSecBufferDesc pOutput, // (inout) Output buffers
'unsigned long SEC_FAR * pfContextAttr  // (out) Context attributes
'PTimeStamp ptsExpiry// (out) Life span (OPT)
Public Function AcceptSecurityContext( _
    ByRef phCredential As DWORD, ByVal phContext As Long, _
    ByRef pInput As SecBufferDesc, ByVal fContextReq As Long, _
    ByVal TargetDataRep As Long, ByRef phNewContext As DWORD, _
    ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, _
    ByRef ptsExpiry As DWORD) As Long
    If IsNT() Then
        AcceptSecurityContext = AcceptSecurityContextNT(phCredential, _
        phContext, pInput, fContextReq, TargetDataRep, phNewContext, _
        pOutput, pfContextAttr, ptsExpiry)
    Else
        AcceptSecurityContext = AcceptSecurityContext9X(phCredential, _
        phContext, pInput, fContextReq, TargetDataRep, phNewContext, _
        pOutput, pfContextAttr, ptsExpiry)
    End If
End Function


'PCtxtHandle phContext, // Context to complete
'PSecBufferDesc pToken// Token to complete
Public Function CompleteAuthToken _
    (ByRef phContext As DWORD, _
    ByRef pToken As SecBufferDesc) As Long
    If IsNT() Then
        CompleteAuthToken = CompleteAuthTokenNT(phContext, pToken)
    Else
        CompleteAuthToken = CompleteAuthToken9X(phContext, pToken)
    End If
End Function


Public Function DeleteSecurityContext(ByRef hCtxt As DWORD) As Long
    If IsNT() Then
        DeleteSecurityContext = DeleteSecurityContextNT(hCtxt)
    Else
        DeleteSecurityContext = DeleteSecurityContext9X(hCtxt)
    End If
End Function


Public Function FreeContextBuffer(ByVal pvContextBuffer As Long) As Long
    If IsNT() Then
        FreeContextBuffer = FreeContextBufferNT(pvContextBuffer)
    Else
        FreeContextBuffer = FreeContextBuffer9X(pvContextBuffer)
    End If
End Function


Public Function FreeCredentialsHandle(ByRef hCred As DWORD) As Long
    If IsNT() Then
        FreeCredentialsHandle = FreeCredentialsHandleNT(hCred)
    Else
        FreeCredentialsHandle = FreeCredentialsHandle9X(hCred)
    End If
End Function


Public Function InitSecurityInterface() As Long
    If IsNT() Then
        InitSecurityInterface = InitSecurityInterfaceNT()
    Else
        InitSecurityInterface = InitSecurityInterface9X()
    End If
End Function


Public Function EnumerateSecurityPackages( _
    ByRef pcPackages As Long, _
    ByRef ppPackageInfo As Long) As Long
    If IsNT() Then
        EnumerateSecurityPackages = _
        EnumerateSecurityPackagesNT(pcPackages, ppPackageInfo)
    Else
        EnumerateSecurityPackages = _
        EnumerateSecurityPackages9X(pcPackages, ppPackageInfo)
    End If
End Function
    

Public Function QuerySecurityPackageInfo( _
    ByVal pszPackageName As String, _
    ByRef ppPackageInfo As Long) As Integer
    If IsNT() Then
        QuerySecurityPackageInfo = _
        QuerySecurityPackageInfoNT(pszPackageName, ppPackageInfo)
    Else
        QuerySecurityPackageInfo = _
        QuerySecurityPackageInfo9X(pszPackageName, ppPackageInfo)
    End If
End Function


Public Function GetMsg(i As Long) As String
    Select Case i
        Case SEC_E_OK
        GetMsg = "OK"
        Case SEC_E_INSUFFICIENT_MEMORY
        GetMsg = "E: INSUFFICIENT_MEMORY"
        Case SEC_E_INVALID_HANDLE
        GetMsg = "E: INVALID_HANDLE"
        Case SEC_E_UNSUPPORTED_FUNCTION
        GetMsg = "E: UNSUPPORTED_FUNCTION"
        Case SEC_E_TARGET_UNKNOWN
        GetMsg = "E: TARGET_UNKNOWN"
        Case SEC_E_INTERNAL_ERROR
        GetMsg = "E: INTERNAL_ERROR"
        Case SEC_E_SECPKG_NOT_FOUND
        GetMsg = "E: SECPKG_NOT_FOUND"
        Case SEC_E_NOT_OWNER
        GetMsg = "E: NOT_OWNER"
        Case SEC_E_CANNOT_INSTALL
        GetMsg = "E: CANNOT_INSTALL"
        Case SEC_E_INVALID_TOKEN
        GetMsg = "E: INVALID_TOKEN"
        Case SEC_E_CANNOT_PACK
        GetMsg = "E: CANNOT_PACK"
        Case SEC_E_QOP_NOT_SUPPORTED
        GetMsg = "E: QOP_NOT_SUPPORTED"
        Case SEC_E_NO_IMPERSONATION
        GetMsg = "E: NO_IMPERSONATION"
        Case SEC_E_LOGON_DENIED
        GetMsg = "E: LOGON_DENIED"
        Case SEC_E_UNKNOWN_CREDENTIALS
        GetMsg = "E: UNKNOWN_CREDENTIALS"
        Case SEC_E_NO_CREDENTIALS
        GetMsg = "E: NO_CREDENTIALS"
        Case SEC_E_MESSAGE_ALTERED
        GetMsg = "E: MESSAGE_ALTERED"
        Case SEC_E_OUT_OF_SEQUENCE
        GetMsg = "E: OUT_OF_SEQUENCE"
        Case SEC_E_NO_AUTHENTICATING_AUTHORITY
        GetMsg = "E: NO_AUTHENTICATING_AUTHORITY"
        Case SEC_I_CONTINUE_NEEDED
        GetMsg = "I: CONTINUE_NEEDED"
        Case SEC_I_COMPLETE_NEEDED
        GetMsg = "I: COMPLETE_NEEDED"
        Case SEC_I_COMPLETE_AND_CONTINUE
        GetMsg = "I: COMPLETE_AND_CONTINUE"
        Case SEC_I_LOCAL_LOGON
        GetMsg = "I: LOCAL_LOGON"
        Case SEC_E_BAD_PKGID
        GetMsg = "E: BAD_PKGID"
        Case SEC_E_CONTEXT_EXPIRED
        GetMsg = "E: CONTEXT_EXPIRED"
        Case SEC_E_INCOMPLETE_MESSAGE
        GetMsg = "E: INCOMPLETE_MESSAGE"
        Case SEC_E_INCOMPLETE_CREDENTIALS
        GetMsg = "E: INCOMPLETE_CREDENTIALS"
        Case SEC_E_BUFFER_TOO_SMALL
        GetMsg = "E: BUFFER_TOO_SMALL"
        Case SEC_I_INCOMPLETE_CREDENTIALS
        GetMsg = "I: INCOMPLETE_CREDENTIALS"
        Case SEC_I_RENEGOTIATE
        GetMsg = "I: RENEGOTIATE"
        Case SEC_E_WRONG_PRINCIPAL
        GetMsg = "E: WRONG_PRINCIPAL"
        Case Else
        GetMsg = "Unknown Error"
    End Select
End Function


Public Function SSPILogonUser(User As String, Password As String, _
    Domain As String, errmsg) As Boolean
        Dim i As Long
        Dim ppkgInfo As Long
        Dim hCred As DWORD
        Dim AuthIdentity As SEC_WINNT_AUTH_IDENTITY
        Dim UserBuf(20) As Byte
        Dim DomainBuf(20) As Byte
        Dim PasswordBuf(20) As Byte
        Dim hCtxt As DWORD
        Dim OutBuffDesc As SecBufferDesc
        Dim OutSecBuff As SecBuffer
        Dim ContextAttributes As Long
        Dim LifeTime As DWORD
        Dim cbMaxMessage As Long
        
        AuthIdentity.Domain = VarPtr(DomainBuf(0))
        AuthIdentity.DomainLength = Len(Domain)
        AuthIdentity.Password = VarPtr(PasswordBuf(0))
        AuthIdentity.PasswordLength = Len(Password)
        AuthIdentity.User = VarPtr(UserBuf(0))
        AuthIdentity.UserLength = Len(User)
        AuthIdentity.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI
        
        StrToByte Domain, DomainBuf
        StrToByte User, UserBuf
        StrToByte Password, PasswordBuf
        i = InitSecurityInterface
        If i < 0 Then GoTo Error
        i = QuerySecurityPackageInfo("NTLM", ppkgInfo)
        If i < 0 Then GoTo Error
        CopyMemory cbMaxMessage, ppkgInfo + 8, 4
        i = FreeContextBuffer(ppkgInfo)
        If i < 0 Then GoTo Error
        '----------------------------------- negotiate
        ReDim pOut(cbMaxMessage) As Byte
        OutSecBuff.cbBuffer = cbMaxMessage
        OutSecBuff.pvBuffer = VarPtr(pOut(0))
        OutSecBuff.BufferType = SECBUFFER_TOKEN
        OutBuffDesc.ulVersion = 0
        OutBuffDesc.cBuffers = 1
        OutBuffDesc.pBuffers = VarPtr(OutSecBuff)
        i = AcquireCredentialsHandle(0, "NTLM", SECPKG_CRED_OUTBOUND, 0, _
        VarPtr(AuthIdentity), 0, 0, hCred, LifeTime)
        If i < 0 Then GoTo Error
        
        i = InitializeSecurityContext(hCred, 0, "AuthSamp", 0, 0, _
        SECURITY_NATIVE_DREP, 0, 0, hCtxt, OutBuffDesc, _
        ContextAttributes, LifeTime)
        If i < 0 Then GoTo Error


        If i = SEC_I_COMPLETE_NEEDED Or i = SEC_I_COMPLETE_AND_CONTINUE Then
            i = CompleteAuthToken(hCtxt, OutBuffDesc)
            MsgBox ("COMPLETE should not be required For NTLM.")
        End If
        '----------------------------------- challenge
        Dim hCred2 As DWORD
        Dim hctxt2 As DWORD
        Dim InBuffDesc2 As SecBufferDesc
        Dim InSecBuff2 As SecBuffer
        Dim OutBuffDesc2 As SecBufferDesc
        Dim OutSecBuff2 As SecBuffer
        ReDim pOut2(cbMaxMessage) As Byte
        i = AcquireCredentialsHandle(0, "NTLM", SECPKG_CRED_INBOUND, 0, _
        0, 0, 0, hCred2, LifeTime)
        If i < 0 Then GoTo Error
        
        InSecBuff2.cbBuffer = OutSecBuff.cbBuffer
        InSecBuff2.pvBuffer = OutSecBuff.pvBuffer
        InSecBuff2.BufferType = SECBUFFER_TOKEN
        InBuffDesc2.ulVersion = 0
        InBuffDesc2.cBuffers = 1
        InBuffDesc2.pBuffers = VarPtr(InSecBuff2)
        OutSecBuff2.cbBuffer = cbMaxMessage
        OutSecBuff2.pvBuffer = VarPtr(pOut2(0))
        OutSecBuff2.BufferType = SECBUFFER_TOKEN
        OutBuffDesc2.ulVersion = 0
        OutBuffDesc2.cBuffers = 1
        OutBuffDesc2.pBuffers = VarPtr(OutSecBuff2)
        i = AcceptSecurityContext(hCred2, 0, InBuffDesc2, 0, SECURITY_NATIVE_DREP, _
        hctxt2, OutBuffDesc2, ContextAttributes, LifeTime)
        If i < 0 Then GoTo Error
        '----------------------------------- authenticate
        Dim InSecBuff As SecBuffer
        Dim InBuffDesc As SecBufferDesc
        InSecBuff.cbBuffer = OutSecBuff2.cbBuffer
        InSecBuff.pvBuffer = OutSecBuff2.pvBuffer
        InSecBuff.BufferType = SECBUFFER_TOKEN
        InBuffDesc.ulVersion = 0
        InBuffDesc.cBuffers = 1
        InBuffDesc.pBuffers = VarPtr(InSecBuff)
        OutSecBuff.cbBuffer = cbMaxMessage
        i = InitializeSecurityContext(hCred, VarPtr(hCtxt), "AuthSamp", 0, 0, _
        SECURITY_NATIVE_DREP, VarPtr(InBuffDesc), 0, hCtxt, OutBuffDesc, _
        ContextAttributes, LifeTime)
        If i < 0 Then GoTo Error
        '----------------------------------- authenticate
        InSecBuff2.cbBuffer = OutSecBuff.cbBuffer
        InSecBuff2.pvBuffer = OutSecBuff.pvBuffer
        OutSecBuff2.cbBuffer = cbMaxMessage
        i = AcceptSecurityContext(hCred2, VarPtr(hctxt2), InBuffDesc2, 0, _
        SECURITY_NATIVE_DREP, hctxt2, OutBuffDesc2, ContextAttributes, LifeTime)
        If i < 0 Then GoTo Error
        i = DeleteSecurityContext(hCtxt)
        If i < 0 Then GoTo Error
        i = DeleteSecurityContext(hctxt2)
        If i < 0 Then GoTo Error
        i = FreeCredentialsHandle(hCred)
        If i < 0 Then GoTo Error
        i = FreeCredentialsHandle(hCred2)
        If i < 0 Then GoTo Error
        SSPILogonUser = True
        Exit Function
Error:
        errmsg = GetMsg(i)
        SSPILogonUser = False
End Function
