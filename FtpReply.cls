VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtpReply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Reply As String

Public Enum ReplyCodes
    repPreliminary = 1
    repComplete = 2
    repIntermediate = 3
    repTransient = 4
    repPermanent = 5
End Enum

Public Enum ReplySubCodes
    repSyntax = 0
    repInformational = 1
    repConnections = 2
    repAuthentication = 3
    repFileSystem = 5
End Enum

'Initialisiert ein Reply Objekt
Friend Property Let Reply(Reply1 As String)
    m_Reply = Left(Reply1, Len(Reply1) - 2)
End Property

'Liefert den gesammten Reply
Public Property Get Reply() As String
    Reply = m_Reply
End Property

'Liefert nur den Status Code zurück
Public Property Get Status() As Integer
    Status = Left(m_Reply, 3)
End Property

'Liefert nur den Message Text zurück
Public Property Get Message() As String
    Message = Trim(Mid(m_Reply, 4))
End Property

'Liefert die erste Zahl aus dem drei Zahlencode zurück
Public Function MainState() As ReplyCodes
    MainState = Int(Status / 100)
End Function

'Liefert die zweite Zahl aus dem drei Zahlencode zurück
Public Function SubState() As ReplySubCodes
    SubState = Int((Status - (Int(Status / 100) * 100)) / 10)
End Function
