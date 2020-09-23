Attribute VB_Name = "basStrings"
Option Explicit

'*************************************************************************
'Sucht einen Sub String in einem String und leifert die Position zurück an
'dem der Sub String gefunden wurde. Suchreihenfolge von Links nach Rechts
'*************************************************************************
Public Function lPos(Str As Variant, SubStr As Variant) As Integer
  Dim i As Long
  
  lPos = 0
  If (Len(Str) = 0) Or (Len(SubStr) = 0) Then
    Exit Function
  End If
  For i = 1 To Len(Str) - Len(SubStr) + 1
    If Mid(Str, i, Len(SubStr)) = SubStr Then
      lPos = i
      Exit Function
    End If
  Next i
End Function


'*************************************************************************
'Sucht einen Sub String in einem String und leifert die Position zurück an
'dem der Sub String gefunden wurde. Suchreihenfolge von Rechts nach Links.
'*************************************************************************
Public Function rPos(Str As Variant, SubStr As Variant) As Integer
  Dim i As Long
  
  rPos = 0
  If (Len(Str) = 0) Or (Len(SubStr) = 0) Then
    Exit Function
  End If
  For i = Len(Str) - Len(SubStr) + 1 To 1 Step -1
    If Mid(Str, i, Len(SubStr)) = SubStr Then
      rPos = i
      Exit Function
    End If
  Next i
End Function


'*************************************************************************
'Konvertiert einen String in ein Byte Array
'*************************************************************************
Public Sub StrToByte(s As String, b() As Byte)
    Dim i As Integer
    For i = 0 To Len(s) - 1
        b(i) = Asc(Mid(s, i + 1, 1))
    Next i
    b(i) = 0
End Sub


'*************************************************************************
'Konvertiert ein Byte Array in einen String
'*************************************************************************
Public Sub ByteToStr(b() As Byte, s As String, ByVal Length As Integer)
    Dim i As Integer
    s = ""
    For i = 0 To Length - 1
        s = s & Chr(b(i))
    Next i
End Sub

'*************************************************************************
'Hilfsmethode für Base64Encode liefert das zugehörige Zeichen zu dem
'6 Bit Zahlencode
'*************************************************************************
Private Function Base64Char(ByVal bit6Number As Byte) As String
    Select Case bit6Number
        Case 0 To 25   'A bis Z
            Base64Char = Chr(65 + bit6Number)
        Case 26 To 51  'a bis z
            Base64Char = Chr(97 + (bit6Number - 26))
        Case 52 To 61  '0 bis 9
            Base64Char = Chr(48 + (bit6Number - 52))
        Case 62        '+
            Base64Char = "+"
        Case 63        '-
            Base64Char = "/"
        Case Else
            MsgBox "Codierungsfehler! bit6Number > 63", vbOKOnly + vbCritical, "Base64Char"
            
            '---------------------------------------
            'Hier noch ein Fehlerbehandlung bla ldsa
            '----------------------------------------
            
    End Select
End Function


'*************************************************************************
'Codiert einen String mit der Base64 Methode
'*************************************************************************
Public Function Base64Encode(ByVal BinaryData As String) As String
    Dim retString As String
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim Byte3 As Byte
    Dim i As Integer
    Dim x As Integer

    Do While Len(BinaryData) > 0
        Byte1 = Asc(BinaryData)
        BinaryData = Mid(BinaryData, 2)
        x = 2
        If Len(BinaryData) >= 1 Then
            Byte2 = Asc(BinaryData)
            BinaryData = Mid(BinaryData, 2)
            x = 1
        End If
        If Len(BinaryData) >= 1 Then
            Byte3 = Asc(BinaryData)
            BinaryData = Mid(BinaryData, 2)
            x = 0
        End If
        
        retString = retString & Base64Char(Int(Byte1 / 4))
        retString = retString & Base64Char(((Byte1 And 3) * 16) + Int(Byte2 / 16))
        retString = retString & Base64Char(((Byte2 And 15) * 4) + Int(Byte3 / 64))
        retString = retString & Base64Char(Byte3 And 63)
    Loop
    If x = 1 Then
        retString = Left(retString, Len(retString) - 1) & "="
    ElseIf x = 2 Then
        retString = Left(retString, Len(retString) - 2) & "=="
    End If
    Base64Encode = retString
End Function


'*************************************************************************
' Decodiert einen String mit der Base64 Methode
'*************************************************************************
Public Function Base64Decode(AsciiData As String) As String
    Dim counter As Integer
    Dim Temp As String
    'For the dec. Tab
    Dim DecodeTable As Variant
    Dim Out(2) As Byte
    Dim inp(3) As Byte
    'DecodeTable holds the decode tab
    DecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", _
    "18", "19", "20", "21", "22", "23", "24", "25", "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255" _
    , "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")
    'Reads 4 Bytes in and decrypt them


    For counter = 1 To Len(AsciiData) Step 4
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '     !!!!!!!!!!!!!!!!!!!
        '!IF YOU WANT YOU CAN ADD AN ERRORCHECK:
        '     !
        '!If DecodeTable()=255 Then Error!!
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '     !!!!!!!!!!!!!!!!!!!
        '4 Bytes in -> 3 Bytes out
        inp(0) = DecodeTable(Asc(Mid$(AsciiData, counter, 1)))
        inp(1) = DecodeTable(Asc(Mid$(AsciiData, counter + 1, 1)))
        inp(2) = DecodeTable(Asc(Mid$(AsciiData, counter + 2, 1)))
        inp(3) = DecodeTable(Asc(Mid$(AsciiData, counter + 3, 1)))
        Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
        Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
        Out(2) = ((inp(2) And &H3) * 64) Or inp(3)
        '* look for "=" symbols


        If inp(2) = 64 Then
            'If there are 2 characters left -> 1
            '     binary out
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Temp = Temp & Chr(Out(0) And &HFF)
        ElseIf inp(3) = 64 Then
            'If there are 3 characters left -> 2
            '     binaries out
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF)
        Else 'Return three Bytes
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF) & Chr(Out(2) And &HFF)
        End If
    Next
    Base64Decode = Temp
End Function





'*************************************************************************
'Einfacher Verschlüsselungs Algortihmus,
'  Codiert die Daten mit dem Key,
'  Codiert die Daten mit Base64Encode,
'  Liefert das Ergebnis zurück
'*************************************************************************
Public Function Encrypt(Message As String, Key As String) As String
    'Variablen deklaration
    Dim i As Integer
    Dim n As Integer
    Dim c As Byte
    Dim retStr As String
    
    'Message verschlüsseln
    For i = 1 To Len(Message)
        n = n + 1
        c = Asc(Mid(Message, i, 1))
        c = c + Asc(Mid(Key, n, 1))
        retStr = retStr & Chr(c)
        If n >= Len(Key) Then n = 0
    Next i
    
    'Chiffre zurückgeben
    Encrypt = Base64Encode(retStr)
End Function


'*************************************************************************
'Einfacher Entschlüsselungs Algorithmus,
'  Decodiert die Chiffre mit Base64Decode,
'  Decodiert die Daten, mit dem Key,
'  Gibt das Ergebnis zurück
'*************************************************************************
Public Function Decrypt(Chiffre As String, Key As String) As String
    'Variablen deklaration
    Dim i As Integer
    Dim n As Integer
    Dim c As Byte
    Dim retStr As String
    Dim Tmp As String
    
    'Base64Decode
    Tmp = Base64Decode(Chiffre)
    
    'Message entschlüsseln
    For i = 1 To Len(Tmp)
        n = n + 1
        c = Asc(Mid(Tmp, i, 1))
        c = c - Asc(Mid(Key, n, 1))
        retStr = retStr & Chr(c)
        If n >= Len(Key) Then n = 0
    Next i
    
    'Chiffre zurückgeben
    Decrypt = retStr
End Function
