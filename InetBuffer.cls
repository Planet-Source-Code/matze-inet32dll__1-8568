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


'Buffer Inhalt zur�cksetzen
Friend Sub Clear()
    m_LengthWritten = 0
    m_Length = 0
    ReDim m_Buffer(0)
End Sub


'Datenblock in dem Buffer hinten anh�ngen
Friend Sub Append(Data As String)
    'Buffer vergr��ern
    m_Length = m_Length + Len(Data)
    m_LengthWritten = m_LengthWritten + Len(Data)
    m_Buffer(UBound(m_Buffer)) = Data
    ReDim Preserve m_Buffer(UBound(m_Buffer) + 1)
End Sub


'Zur�ckgeben wieviel Daten schon geschrieben wurden
Friend Property Get Written() As Long
    Written = m_LengthWritten
End Property


'Datenblock in den Buffer einf�gen
'Friend Sub AddData(Data As String, Optional Offset As Long)
    'Variablen deklaration
'    Dim Segment As Long
'    Dim tmpOffset As Long
'    Dim c As Long
'    Dim i As Long
    
    'Pr�fen ob das Objekt Schreibgesch�tzt ist
'    If m_Protected Then Err.Raise inetProtected, "Inet32.InetBuffer", "Dieses Objekt ist Schreibgesch�tzt"
    
    'Pr�fen ob die Buffer Grenzen �berschritten werden
'    tmpOffset = Offset
'    If tmpOffset < 0 Or tmpOffset > m_Length Then
'        Err.Raise vbObjectError, "Inet32.InetResponse", "Buffer Grenzen �berschritten."
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
 '       'Buffer vergr��ern und Daten neu anordnen
 '       ReDim Preserve m_Buffer(UBound(m_Buffer) + 1)
 '       For i = UBound(m_Buffer) - 1 To Segment + 1 Step -1
 '           m_Buffer(i) = m_Buffer(i - 1)
 '       Next i
 '       'Daten in den Buffer kopieren
 '       m_Buffer(Segment) = Data
 '   Else
 '       'Buffer vergr��ern und Daten neu anordnen
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
 '   'L�ngen Variable anpassen
 '   m_Length = m_Length + Len(Data)
'End Sub


'Datenblock zur�ckgeben
Public Sub PeekData(Data As String, Optional Length As Long)
    'Variablen deklaration
    Dim Segment As Long
    Dim tmpLength As Long
    
    'Pr�fen ob die Buffer Grenzen �berschritten werden
    If Length > m_Length Then
        Err.Raise vbObjectError, "Inet32.InetResponse", "Buffer Grenzen �berschritten."
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


'Daten auf dem Buffer holen und l�schen
Public Sub GetData(Data As String, Optional Length As Long)
    'Variablen deklaration
    Dim tmpLength As Long
    Dim curLength As Long
    Dim i As Long
    
    'Daten aus dem Buffer lesen
    PeekData Data, Length
    If Data = "" Then Exit Sub

    'Geholte Daten l�schen
    tmpLength = Len(Data)
    Do While tmpLength > curLength
        If curLength + Len(m_Buffer(i)) > tmpLength Then
            'Nur ein bi�chen l�schen
            m_Buffer(i) = Mid(m_Buffer(i), tmpLength - curLength + 1)
            curLength = tmpLength
        Else
            'Gesammten String l�schen
            curLength = curLength + Len(m_Buffer(i))
            m_Buffer(i) = ""
        End If
        'Index erh�hen
        i = i + 1
    Loop
    
    'L�ngen Variable anpassen
    m_Length = m_Length - tmpLength
End Sub


'Genau die selbe Funktion wie PeekData nur als Eigenschaft
'und man kann die L�nge nicht bestimmen
Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
    'Variablen deklaration
    Dim tmp As String
    PeekData tmp    'Daten auslesen
    Data = tmp      'Zur�ckgeben
End Property


'Bufferl�nge zur�ck geben
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
