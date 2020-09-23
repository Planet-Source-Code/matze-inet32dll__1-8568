VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents x As HttpEngine
Attribute x.VB_VarHelpID = -1
Dim y As FtpEngine


Private Sub FileTest()
    Dim z As New fileengine
    Dim y As InetBuffer
    Dim t As String
    Set y = z.OpenURL("file://c:\autoexec.bat", True)
    
    Debug.Print y.Length
    y.GetData t, 10
    Debug.Print t
    Debug.Print y.Length
    
    Debug.Print y.Length
    y.GetData t, 10
    Debug.Print t
    Debug.Print y.Length

End Sub



Private Sub Form_Load()
    'Dim z As New HttpEngine
    'z.Proxy.Access = inetDefault
    'z.Proxy.Server = "192.168.1.1"
    'z.Proxy.Port = "4480"
    
    'Debug.Print z.Proxy.Server
    'Debug.Print z.Proxy.Port
    
    'Debug.Print z.OpenURL("http://192.168.1.3/index.htm", True, "GET")
    
    Dim y As New FtpEngine
    
    'y.Proxy.Access = inetDefault
    'y.Proxy.Server = "localhost:4480"
    'y.Passive = True
    'Debug.Print y.Exist("ftp://192.168.1.3/security.htm")
    'Debug.Print y.Exist("ftp://192.168.1.3/indx.htm")
    'Debug.Print y.Exist("ftp://192.168.1.3/index.htm")
    'y.Passive = True
    Debug.Print "connect="; y.Connect("192.168.1.3", 21)
    Debug.Print "login="; y.Login("tux99", "tux99")
    Debug.Print "upload = " & y.Store("c:\autoexec.bat", "/home/tux99/test1.txt")
End Sub



Private Sub HttpTest()
    Set x = New HttpEngine

    x.OpenURL

    x.Proxy.Access = inetnamedproxy
    'x.Proxy.Access = inetDirect
    x.Proxy.Protocol = "HTTP"
    x.Proxy.Server = "ezn.telekom.de"
    x.Proxy.Port = 80
    x.Proxy.Username = "Luebben.Matthias"
    x.Proxy.Password = "habkeins"
    
    x.Proxy.Domain = "EZNORD"
    x.Proxy.Override = "*.telekom.de; w8r00924"
    
    Dim test4 As InetBuffer
    Dim s As String
    
    'x.OpenURL "http://www.t-online.de/nix.htm", False, "GET"
    'x.OpenURL "http://w8r00924/Deadlink/redirect2tonline.htm", False, "GET"
    'x.OpenURL "http://w8r00924/Deadlink/writeable.htm", False, "POST", "Hallo Welt"
    'x.OpenURL "http://nix.t-online.de", True, "GET"
    
    x.OpenURL "http://home.t-online.de/home/Christian.Gellert", True, "GET"
    'Dim y As String
    'x.Buffer.GetData y
    'Debug.Print y
    
    Debug.Print x.Header.Resource
    x.Buffer.GetData s
    Debug.Print s
    
    'Debug.Print "Site State: " & x.Exist("http://www.t-online.de/index.htm")
End Sub



Private Sub x_Challenge()
    Debug.Print "Challange"
End Sub

Private Sub x_Complete()
    Debug.Print x.Header.Status
    Dim tmp As String
    x.Buffer.GetData tmp
    Debug.Print tmp
End Sub

Private Sub x_Connected()
    Debug.Print "Connected"
End Sub

Private Sub x_Connecting(ByVal Proxy As String, ByVal Host As String)
    Debug.Print "Connecting: " & Proxy & " / " & Host
End Sub

Private Sub x_Error(ByVal Number As Long, ByVal Reason As String)
    Debug.Print "Error: " & Number & " / " & Reason
End Sub

Private Sub x_HeaderReceived()
    Debug.Print "HeaderReceived"
End Sub

Private Sub x_Progress(ByVal bytesReceived As Long, ByVal bytesTotal As Long)
    Debug.Print "Progess: " & bytesReceived & " / " & bytesTotal
End Sub

Private Sub x_Redirection(ByVal FromUrl As Inet32.InetURL, ToUrl As Inet32.InetURL)
    Debug.Print "Redirection: " & FromUrl & " / " & ToUrl
End Sub

Private Sub x_Requesting(ByVal methode As String, ByVal Resource As Inet32.InetURL)
    Debug.Print "Requesting: " & methode & " / " & Resource
End Sub
