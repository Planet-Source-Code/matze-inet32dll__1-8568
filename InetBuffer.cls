VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InetBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Data Members
Private m_LengthWritten As Long
Private m_Length As Long
Private m_Buffer() As String


'Buffer Inhalt zurücksetzen
Friend Sub Clear()
    m_LengthWritten = 0
    m_Length = 0
    ReDim m_Buffer(0)
End Sub


'Datenblock in dem Buffer hinten anhängen
Friend Sub Append(Data As String)
    'Buffer vergrößern
    m_Length = m_Length + Len(Data)
    m_LengthWritten = m_LengthWritten + Len(Data)
    m_Buffer(UBound(m_Buffer)) = Data
    ReDim Preserve m_Buffer(UBound(m_Buffer) + 1)
End Sub


'Zurückgeben wieviel Daten schon geschrieben wurden
Friend Property Get Written() As Long
    Written = m_LengthWritten
End Property


'Datenblock in den Buffer einfügen
'Friend Sub AddData(Data As String, Optional Offset As Long)
    'Variablen deklaration
'    Dim Segment As Long
'    Dim tmpOffset As Long
'    Dim c As Long
'    Dim i As Long
    
    'Prüfen ob das Objekt Schreibgeschützt ist
'    If m_Protected Then Err.Raise inetProtected, "Inet32.InetBuffer", "Dieses Objekt ist Schreibgeschützt"
    
    'Prüfen ob die Buffer Grenzen überschritten werden
'    tmpOffset = Offset
'    If tmpOffset < 0 Or tmpOffset > m_Length Then
'        Err.Raise vbObjectError, "Inet32.InetResponse", "Buffer Grenzen überschritten."
'    End If
'    If tmpOffset = 0 Then tmpOffset = 1
        
    'Offset suchen
 '   Do While c < tmpOffset
 '       c = c + Len(m_Buffer(Segment))
 '       Segment = Segment + 1
 '   Loop
 '   Segment = Segment - 1
 '   c = c - Len(m_Buffer(Segment))
 '   c = tmpOffset - c
 '
 '   If c = 1 Then
 '       'Buffer vergrößern und Daten neu anordnen
 '       ReDim Preserve m_Buffer(UBound(m_Buffer) + 1)
 '       For i = UBound(m_Buffer) - 1 To Segment + 1 Step -1
 '           m_Buffer(i) = m_Buffer(i - 1)
 '       Next i
 '       'Daten in den Buffer kopieren
 '       m_Buffer(Segment) = Data
 '   Else
 '       'Buffer vergrößern und Daten neu anordnen
 '       ReDim Preserve m_Buffer(UBound(m_Buffer) + 2)
 '       For i = UBound(m_Buffer) - 1 To Segment + 2 Step -1
 '           m_Buffer(i) = m_Buffer(i - 2)
 '       Next i
 '       m_Buffer(Segment + 2) = Mid(m_Buffer(Segment), c)
 '       m_Buffer(Segment) = Left(m_Buffer(Segment), c - 1)
 '
 '       'Daten in den Buffer kopieren
 '       m_Buffer(Segment + 1) = Data
 '   End If
 '
 '   'Längen Variable anpassen
 '   m_Length = m_Length + Len(Data)
'End Sub


'Datenblock zurückgeben
Public Sub PeekData(Data As String, Optional Length As Long)
    'Variablen deklaration
    Dim Segment As Long
    Dim tmpLength As Long
    
    'Prüfen ob die Buffer Grenzen überschritten werden
    If Length > m_Length Then
        Err.Raise vbObjectError, "Inet32.InetResponse", "Buffer Grenzen überschritten."
    End If
    
    'Ausgabe Buffer vorbereiten
    If Length > 0 Then
        tmpLength = Length
    Else
        tmpLength = m_Length
    End If
    
    If tmpLength > 0 Then
        'Daten in den Buffer kopieren
        Data = ""
        Do While Len(Data) < tmpLength
            Data = Data + m_Buffer(Segment)
            Segment = Segment + 1
        Loop
        Data = Left(Data, tmpLength)
    Else
        Data = ""
    End If
End Sub


'Daten auf dem Buffer holen und löschen
Public Sub GetData(Data As String, Optional Length As Long)
    'Variablen deklaration
    Dim tmpLength As Long
    Dim curLength As Long
    Dim i As Long
    
    'Daten aus dem Buffer lesen
    PeekData Data, Length
    If Data = "" Then Exit Sub

    'Geholte Daten löschen
    tmpLength = Len(Data)
    Do While tmpLength > curLength
        If curLength + Len(m_Buffer(i)) > tmpLength Then
            'Nur ein bißchen löschen
            m_Buffer(i) = Mid(m_Buffer(i), tmpLength - curLength + 1)
            curLength = tmpLength
        Else
            'Gesammten String löschen
            curLength = curLength + Len(m_Buffer(i))
            m_Buffer(i) = ""
        End If
        'Index erhöhen
        i = i + 1
    Loop
    
    'Längen Variable anpassen
    m_Length = m_Length - tmpLength
End Sub


'Genau die selbe Funktion wie PeekData nur als Eigenschaft
'und man kann die Länge nicht bestimmen
Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
    'Variablen deklaration
    Dim tmp As String
    PeekData tmp    'Daten auslesen
    Data = tmp      'Zurückgeben
End Property


'Bufferlänge zurück geben
Public Property Get Length() As Long
    Length = m_Length
End Property


'Speichern in Datei
Public Sub SaveFile(FileName As String)
    'Variablen deklaration
    Dim ff As Integer
    Dim i As Long
    
    'Speichern in Datei
    ff = FreeFile
    Open FileName For Binary Access Write As #ff
    For i = 0 To UBound(m_Buffer) - 1
        Put #ff, , m_Buffer(i)
    Next i
    Close #ff
End Sub


'Konstruktion
Private Sub Class_Initialize()
    ReDim m_Buffer(0)
End Sub
