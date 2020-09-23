VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**************************************

' Name: cReadEasyReg (Updated)

' Description:I normally only want to re

'     ad the registry and found all the other

'     classes to difficult to use, so I wrote

'     my own. This one is a update on the prev

'     ious published cReadEasyReg on this site

'     . This class will also read from the NT

'     Registry :> PLEASE NOTE: This is a cl

'     ass module and all the code should be pa

'     ste into a CLASS Module.

' By: Riaan Aspeling

'

'

' Inputs:None

'

' Returns:Returns values from the regist

'     ry.

'

'Assumes:Like any registry code.. (even

'     if this one only reads information) you

'     should be cautious using it.

'

'Side Effects:Some types of values are n

'     ot supported but I haven't encountered t

'     hem in the registry jet.

'

'Warranty:

'Code provided by Planet Source Code(tm)

'     (http://www.Planet-Source-Code.com) 'as

'     is', without warranties as to performanc

'     e, fitness, merchantability,and any othe

'     r warranty (whether expressed or implied

'     ).

'Terms of Agreement:

'By using this source code, you agree to

'     the following terms...

' 1) You may use this source code in per

'     sonal projects and may compile it into a

'     n .exe/.dll/.ocx and distribute it in bi

'     nary format freely and with no charge.

' 2) You MAY NOT redistribute this sourc

'     e code (for example to a web site) witho

'     ut written permission from the original

'     author.Failure to do so is a violation o

'     f copyright laws.

' 3) You may link to this code from anot

'     her website, provided it is not wrapped

'     in a frame.

' 4) The author of this code may have re

'     tained certain additional copyright righ

'     ts.If so, this is indicated in the autho

'     r's description.

'**************************************


' Developed by : Riaan Aspeling

' Company :Altered Reality Corporation

' Date : 1999-Apr-10

' Country :South Africa

'

' Description : A Easy way to READ the r

'     egistry

' Comment :Most of the times a work with

'     the registry is only want

' to READ it, not write to it. Hope you

'     guys/gals out there

' could use this code.

' Problems :If you do find any problems

'     (not Microsoft related) let me

' know at :

'arc@iti.co.za

' Have fun reading the registry ;-)

Option Explicit
'Registry API's to use



Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long


Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long


Private Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
    'Enum's for the OpenRegistry function



Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum


'Registry Type's

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
'Right's for the OpenRegistry

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = ( _
KEY_READ _
And _
KEY_WRITE _
)
Private Const KEY_ALL_ACCESS = ( _
( _
STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or _
KEY_SET_VALUE Or _
KEY_CREATE_SUB_KEY Or _
KEY_ENUMERATE_SUB_KEYS Or _
KEY_NOTIFY Or _
KEY_CREATE_LINK _
) _
And _
( _
Not SYNCHRONIZE _
) _
)
'Local var's to keep track of things hap

'     pening

Dim RootHKey As HKeys
Dim SubDir As String
Dim HKey As Long
Dim OpenRegOk As Boolean
'This function will return a array of va

'     riant with all the subkey values

'eg.

'Dim MyVariant As Variant, MyReg As New

'     CReadEasyReg, i As Integer

'If Not MyReg.OpenRegistry(HKEY_LOCAL_MA

'     CHINE, "Software\Microsoft") Then

' MsgBox "Couldn't open the registry"

' Exit Sub

'End If

'MyVariant = MyReg.GetAllSubDirectories

'For i = LBound(MyVariant) To UBound(MyV

'     ariant)

' Debug.Print MyVariant(i)

'Next i

'MyReg.CloseRegistry



Function GetAllSubDirectories() As Variant


    On Error GoTo handelgetdirvalues
    Dim SubKey_Num As Integer
    Dim SubKey_Name As String
    Dim Length As Long
    Dim ReturnArray() As Variant
    
    If Not OpenRegOk Then Exit Function
    'Get the Dir List



SubKey_Num = 0




    Do
        Length = 256


SubKey_Name = Space$(Length)




    If RegEnumKey(HKey, SubKey_Num, SubKey_Name, Length) <> 0 Then
        Exit Do
    End If




SubKey_Name = Left$(SubKey_Name, InStr(SubKey_Name, Chr$(0)) - 1)


    ReDim Preserve ReturnArray(SubKey_Num) As Variant
    ReturnArray(SubKey_Num) = SubKey_Name


SubKey_Num = SubKey_Num + 1


Loop


GetAllSubDirectories = ReturnArray
Exit Function
handelgetdirvalues:
GetAllSubDirectories = Null
Exit Function
End Function


'This function will return a array of va

'     riant with all the value names in a key

'eg.

'Dim MyVariant As Variant, MyReg As New

'     CReadEasyReg, i As Integer

'If Not MyReg.OpenRegistry(HKEY_LOCAL_MA

'     CHINE, "HardWare\Description\System\Cent

'     ralProcessor\0") Then

' MsgBox "Couldn't open the registry"

' Exit Sub

'End If

'MyVariant = MyReg.GetAllValues

'For i = LBound(MyVariant) To UBound(MyV

'     ariant)

' Debug.Print MyVariant(i)

'Next i

'MyReg.CloseRegistry



Function GetAllValues() As Variant


    On Error GoTo handelgetdirvalues
    Dim lpData As String, KeyType As Long
    Dim BufferLengh As Long, vname As String, vnamel As Long
    Dim ReturnArray() As Variant, Index As Integer
    
    If Not OpenRegOk Then Exit Function
    
    'Get the Values List

    Index = 0


    Do
        lpData = String(250, " ")
        BufferLengh = 240
        vname = String(250, " ")
        vnamel = 240


        If RegEnumValue(ByVal HKey, ByVal Index, vname, vnamel, 0, KeyType, lpData, BufferLengh) <> 0 Then
            Exit Do
        End If


        vname = Left$(vname, InStr(vname, Chr$(0)) - 1)
        ReDim Preserve ReturnArray(Index) As Variant
        ReturnArray(Index) = vname
        Index = Index + 1
    Loop


    GetAllValues = ReturnArray
    Exit Function
handelgetdirvalues:
    GetAllValues = Null
    Exit Function
End Function


'This function will return a specific va

'     lue from the registry

'eg.

'Dim MyString As String, MyReg As New CR

'     eadEasyReg, i As Integer

'If Not MyReg.OpenRegistry(HKEY_LOCAL_MA

'     CHINE, "HardWare\Description\System\Cent

'     ralProcessor\0") Then

' MsgBox "Couldn't open the registry"

' Exit Sub

'End If

'MyString = MyReg.GetValue("Identifier")

'

'Debug.Print MyString

'MyReg.CloseRegistry



Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As String
    On Error GoTo handelgetavalue
    Dim i As Integer
    Dim SubKey_Value As String, TempStr As String
    Dim Length As Long
    Dim value_type As Long, RtnVal As Long
    
    If Not OpenRegOk Then Exit Function
    
    'Read the size of the value value
    RtnVal = RegQueryValueEx(HKey, VarName, 0&, value_type, ByVal 0&, Length)


    Select Case RtnVal
        Case 0 'Ok so continue
        Case 2 'Not Found
        Exit Function
        Case 5 'Access Denied
        GetValue = "Access Denied"
        Exit Function
        Case Else 'What?
        GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
        Exit Function
    End Select


'declare the size of the value and read

'     it

    SubKey_Value = Space$(Length)
    RtnVal = RegQueryValueEx(HKey, VarName, 0&, value_type, ByVal SubKey_Value, Length)

    Select Case value_type
        Case REG_NONE
            'Not defined
            SubKey_Value = "Not defined value_type=REG_NONE"

        Case REG_SZ
            'A null-terminated String
            SubKey_Value = Left$(SubKey_Value, Length - 1)

        Case REG_EXPAND_SZ
            'A null-terminated string that contains
            'unexpanded references to
            'environment variables (for example, "%P
            'ATH%").
            'Use ExpandEnvironmentStrings to expand
            SubKey_Value = Left$(SubKey_Value, Length - 1)

        Case REG_BINARY
            'Binary data in any form.
            SubKey_Value = Left$(SubKey_Value, Length)

    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i




SubKey_Value = TempStr


End If


Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN 'A 32-bit number.


SubKey_Value = Left$(SubKey_Value, Length)




    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i




SubKey_Value = TempStr


End If


Case REG_DWORD_BIG_ENDIAN
'A 32-bit number in big-endian format.

'In big-endian format, a multi-byte valu

'     e is stored in memory from

'the highest byte (the "big end") to the

'     lowest byte. For example,

'the value 0x12345678 is stored as (0x12

'     0x34 0x56 0x78) in big-endian

'format.

Case REG_LINK
'A Unicode symbolic link. Used internall

'     y; applications should not

'use this type.



SubKey_Value = "Not defined value_type=REG_LINK"


    Case REG_MULTI_SZ
    'Array of null-terminated string



SubKey_Value = Left$(SubKey_Value, Length)


    Case REG_RESOURCE_LIST
    'Device driver resource list.



SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"


    Case REG_FULL_RESOURCE_DESCRIPTOR
    'Device driver resource list.



SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"


    Case REG_RESOURCE_REQUIREMENTS_LIST
    'Device driver resource list.



SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"


    Case Else


SubKey_Value = "value_type=" & value_type


End Select


GetValue = SubKey_Value
Exit Function
handelgetavalue:
GetValue = ""
Exit Function
End Function


'This property returns the current KeyVa

'     lue



Public Property Get RegistryRootKey() As HKeys


    RegistryRootKey = RootHKey
End Property


'This property returns the current 'Regi

'     stry Directory' your in



Public Property Get SubDirectory() As String




SubDirectory = SubDir


End Property


'This function open's the registry at a

'     specific 'Registry Directory'

'eg.

'Dim MyVariant As Variant, MyReg As New

'     CReadEasyReg, i As Integer

'If Not MyReg.OpenRegistry(HKEY_LOCAL_MA

'     CHINE, "") Then

' MsgBox "Couldn't open the registry"

' Exit Sub

'End If

'MyVariant = MyReg.GetAllSubDirectories

'For i = LBound(MyVariant) To UBound(MyV

'     ariant)

' Debug.Print MyVariant(i)

'Next i

'MyReg.CloseRegistry



Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Integer


    On Error GoTo OpenReg
    Dim ReturnVal As Integer


    If RtHKey = 0 Then
        OpenRegistry = False
        OpenRegOk = False
        Exit Function
    End If


    RootHKey = RtHKey


SubDir = SbDr




    If OpenRegOk Then
        CloseRegistry
        OpenRegOk = False
    End If


    ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_READ, HKey)


    If ReturnVal <> 0 Then
        OpenRegistry = ReturnVal
        Exit Function
    End If


    OpenRegOk = True
    OpenRegistry = True
    Exit Function
OpenReg:
    OpenRegOk = False
    OpenRegistry = False
    Exit Function
End Function


'This function should be called after yo

'     u're done with the registry

'eg. (see other examples)



Public Function CloseRegistry() As Boolean


    On Error Resume Next


    If RegCloseKey(HKey) <> 0 Then
        CloseRegistry = False
        Exit Function
    End If


    CloseRegistry = True
    OpenRegOk = False
End Function




Private Sub Class_Initialize()


    RootHKey = &H0


SubDir = ""


    HKey = 0
    OpenRegOk = False
End Sub




Private Sub Class_Terminate()
    On Error Resume Next
    If RegCloseKey(HKey) <> 0 Then
        Exit Sub
    End If
End Sub

