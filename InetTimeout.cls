VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InetTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Data Members
Private m_Protected As Boolean
Private m_Connection As Long
Private m_Data As Long

'GET Methoden
Public Property Get Connection() As Long: Connection = m_Connection: End Property
Public Property Get Data() As Long: Data = m_Data: End Property

'LET Methoden
Public Property Let Connection(Milliseconds As Long): m_Connection = Milliseconds: End Property
Public Property Let Data(Milliseconds As Long): m_Data = Milliseconds: End Property

'Friend Methoden
Friend Property Get Protected() As Boolean: Protected = m_Protected: End Property
Friend Property Let Protected(NewProtectedState As Boolean): m_Protected = NewProtectedState: End Property

'Initialisierung
Private Sub Class_Initialize()
    m_Connection = 10000
    m_Data = 20000
End Sub
