VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Inet32.DLL Example ""HTTPEngine"""
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Request file !"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "http://"
      Top             =   300
      Width           =   4575
   End
   Begin VB.ListBox List1 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Dateiname, in dem die Anforderung gespeichert werden soll:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Anzufordernder URL:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   3720
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim a            As Long
Dim actHost      As String
Dim actPort      As String
Dim actProgress  As Long
Dim URLToRequest As String
Dim Dateiname    As String
Dim StartTime    As Single
Dim ViaTime      As Single
Dim EndTime      As Single
Dim DauerInSec   As Single
Dim BytesSec     As Single

Dim WithEvents HTTPRequester As Inet32.HttpEngine
Attribute HTTPRequester.VB_VarHelpID = -1

Private Sub Command1_Click()
On Error GoTo fehler

List1.Clear

URLToRequest = Text1.Text
URLToRequest = Replace(URLToRequest, " ", "%20")
Text1.Text = URLToRequest
Text1.SelStart = Len(Text1.Text)

Set HTTPRequester = New Inet32.HttpEngine
StartTime = (GetTickCount / 1000)
HTTPRequester.OpenURL (URLToRequest)

Exit Sub

fehler:
If Err = -2147221352 Then MsgBox "Ungültiger URL !", vbOKOnly, "Fehler": Text1.SetFocus: Resume Next: Exit Sub
End Sub

Private Sub HTTPRequester_Complete()
On Error GoTo fehler2
EndTime = (GetTickCount / 1000)
DauerInSec = EndTime - StartTime
BytesSec = actProgress / DauerInSec

List1.AddItem "Anforderung komplett"
List1.AddItem "->" & actProgress & " Bytes"
List1.AddItem "->" & FormatNumber(DauerInSec, 2) & " Sekunden"
List1.AddItem "->" & FormatNumber(BytesSec, 0) & " Bytes/Sekunde"
lblProgress.Caption = ""
If Text2.Text = "dev0" Then Exit Sub
Open Text2.Text For Output As #1
  Print #1, HTTPRequester.Buffer
Close
Exit Sub
fehler2:
If Err = 75 Then
  Text2.Text = InputBox("Ungültiger Dateiname !" & vbCrLf & vbCrLf & "Bitte anderen Dateinamen eingeben !", "Fehler !")
  Open Text2.Text For Output As #1
    Print #1, HTTPRequester.Buffer
  Close
  Exit Sub
End If
End Sub

Private Sub HTTPRequester_Connected()
List1.AddItem "Verbunden mit " & actHost & " Port " & actPort
End Sub

Private Sub HTTPRequester_Connecting(ByVal Proxy As String, ByVal Host As String)
actHost = Host
actHost = Left(Host, InStr(1, Host, ":") - 1)
actPort = Mid(Host, InStr(1, Host, ":") + 1, Len(Host))
List1.AddItem "Verbinden mit " & actHost & " ..."
End Sub

Private Sub HTTPRequester_Error(ByVal Number As Long, ByVal Reason As String)
a = MsgBox(Reason, vbOKOnly, "Fehler")
End Sub

Private Sub HTTPRequester_Progress(ByVal bytesReceived As Long, ByVal bytesTotal As Long)
If bytesTotal <> -1 Then
  lblProgress.Caption = "Fortschritt: " & CStr(bytesReceived) & " von " & CStr(bytesTotal) & " Bytes (" & CStr(CInt((bytesReceived / bytesTotal) * 100)) & " %)"
Else
  lblProgress.Caption = "Fortschritt: " & CStr(bytesReceived) & " Bytes"
End If
actProgress = bytesReceived
lblProgress.Refresh
End Sub

Private Sub HTTPRequester_Redirection(ByVal FromUrl As Inet32.InetURL, ToUrl As Inet32.InetURL)
List1.AddItem "Weitergeleitet zu " & ToUrl & "."
End Sub

Private Sub HTTPRequester_Requesting(ByVal Methode As String, ByVal Resource As Inet32.InetURL)
List1.AddItem "Dokument anfordern ..."
End Sub
