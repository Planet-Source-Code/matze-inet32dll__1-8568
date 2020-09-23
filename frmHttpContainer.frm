VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHttpContainer 
   Caption         =   "HTTP Container"
   ClientHeight    =   690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   690
   ScaleWidth      =   2685
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrData 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   600
      Top             =   120
   End
   Begin MSWinsockLib.Winsock sckHttp 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmHttpContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
