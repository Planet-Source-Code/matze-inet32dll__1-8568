VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Data Members
Private m_Resource As New InetURL
Private m_Status As ProtocolStates
Private m_Length As Long
Private m_Request As Date
Private m_Modified As Date
Private m_Attributes As Long

Public Property Get Status() As ProtocolStates
    Status = m_Status
End Property

Friend Property Let Status(NewStatus As ProtocolStates)
    m_Status = NewStatus
End Property

Public Property Get Message() As String
    Select Case m_Status
        Case fileOK: Message = "OK"
        Case fileNotFound: Message = "File Not Found"
        Case fileCanceled: Message = "Canceled by user"
        Case Else: Message = "Unknown State"
    End Select
End Property


Public Property Get Length() As Long
    Length = m_Length
End Property

Friend Property Let Length(NewLength As Long)
    m_Length = NewLength
End Property

Public Property Get Request() As Date
    Request = m_Request
End Property

Friend Property Let Request(NewDate As Date)
    m_Request = NewDate
End Property

Public Property Get Modified() As Date
    Modified = m_Modified
End Property

Friend Property Let Modified(NewDate As Date)
    m_Modified = NewDate
End Property

Public Property Get Attributes() As Long
    Attributes = m_Attributes
End Property

Friend Property Let Attributes(NewAttributes As Long)
    m_Attributes = NewAttributes
End Property

Public Function IsNormal() As Boolean
    IsNormal = m_Attributes And vbNormal
End Function

Public Function IsReadonly() As Boolean
    IsReadonly = m_Attributes And vbReadOnly
End Function

Public Function IsHidden() As Boolean
    IsHidden = m_Attributes And vbHidden
End Function

Public Function IsSystem() As Boolean
    IsSystem = m_Attributes And vbSystem
End Function

Public Function IsArchive() As Boolean
    IsArchive = m_Attributes And vbArchive
End Function
