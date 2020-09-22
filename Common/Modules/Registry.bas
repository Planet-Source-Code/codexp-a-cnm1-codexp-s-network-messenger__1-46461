Attribute VB_Name = "Registry"
'wird in der Form benötigt:
Public SubKey$
Public Eintrag$
Public Key%


'Fehler oder nicht
Const ERROR_SUCCESS = 0

'Hauptschlüssel
Const MainKey = &H80000000
Public Const HKEY_CLASSES_ROOT = 0
Public Const HKEY_CURRENT_USER = 1
Public Const HKEY_LOCAL_MACHINE = 2
Public Const HKEY_USERS = 3
Public Const HKEY_PERFORMANCE_DATA = 4 '(nur NT)
Public Const HKEY_CURRENT_CONFIG = 5
Public Const HKEY_DYN_DATA = 6


'Zugriffcodes
Const KEY_ALL_ACCESS = &H3F

Public Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8

'Optionen beim Anlegen
Const REG_OPTION_NON_VOLATILE = 0

'Datentypen
Const REG_NONE = 0
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_DWORD_LITTLE_ENDIAN = 4
Const REG_DWORD_BIG_ENDIAN = 5
Const REG_LINK = 6
Const REG_MULTI_SZ = 7
Const REG_RESOURCE_LIST = 8
Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Const REG_RESOURCE_REQUIREMENTS_LIST = 10


'Strukturen
Type Time
   LowTime As Long
   HighTime As Long
End Type


'Sicherheitsstruktur nur zur Deklaration der Funktionen
Type SECURITY_ATTRIBUTES
   Length As Long
   Descriptor As Long
   InheritHandle As Boolean
End Type

'Prototypen aus ADVAPI32
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal sSubKey As String, ByVal lReserved As Long, ByVal lSecurity As Long, hKeyReturn As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Declare Function RegQueryValue Lib "advapi32" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal sValue As String, lReserved As Long, lTyp&, ByVal sData As String, lcbData As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpname As String, ByVal cbName As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal SubKeyIndex As Long, ByVal SubKeyName As String, SubKeyNameSize As Long, Reserved As Long, ByVal Class As String, ClassSize As Long, FileTime As Time) As Long

Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal Name As String, ByVal Reserved As Long, ByVal DataType As Long, ByVal Data As String, ByVal Length As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal SubKey As String, ByVal Reserved As Long, ByVal Class As String, ByVal Options As Long, ByVal Access As Long, Security As SECURITY_ATTRIBUTES, hKeyNew As Long, Disposition As Long) As Long
Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal Key As String) As Long
Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Public Sub SaveINI(ByVal Filename As String, ByVal Key As String, ByVal Setting As String, ByVal Value As String)
  If Setting = Chr(0) Then
    Call WritePrivateProfileString(Key, Nothing, Nothing, Filename)
  Else
    If Value = Chr(0) Then
      Call WritePrivateProfileString(Key, Setting, Nothing, Filename)
    Else
      Call WritePrivateProfileString(Key, Setting, Value, Filename)
    End If
  End If
End Sub

Public Function GetINI(ByVal Filename As String, ByVal Key As String, ByVal Setting As String, Optional ByVal Default As String) As String
   Dim Temp As String * 1024
   Call GetPrivateProfileString(Key, Setting, Default, Temp, Len(Temp), Filename)
   GetINI = Mid(Temp, 1, InStr(1, Temp, Chr(0)) - 1)
End Function

Sub RegCreate(MainKey%, ByVal Key$)
    b$ = ""
    Do While InStr(Key$, "\")
        c$ = zeichennext$(Key$, "\")
        RegCreateKey MainKey%, b$, c$
        If Len(b$) Then b$ = b$ + "\"
        b$ = b$ + c$
    Loop
    RegCreateKey MainKey%, b$, Key$
End Sub

Sub RegSetValue(KeyIndex%, SubKey As String, Name As String, lTyp&, Wert As String, lByte&)
'KeyIndex=0: HKEY_CLASSES_ROOT
'         1: HKEY_CURRENT_USER
'         2: HKEY_LOCAL_MACHINE
'         3: HKEY_USERS
'         4: HKEY_PERFORMANCE_DATA (nur NT)
'         5: HKEY_CURRENT_CONFIG
'         6: HKEY_DYN_DATA

    lhKey& = MainKey + KeyIndex
    lResult& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_SET_VALUE, lhKeyOpen&)
    If lResult& <> ERROR_SUCCESS Then Exit Sub
    lResult& = RegSetValueEx(lhKeyOpen&, Name, 0, lTyp&, Wert, lByte&)
    'If lResult& <> ERROR_SUCCESS Then (Fehler...)
    RegCloseKey lhKeyOpen&
End Sub

Sub Reg_DeleteValue(KeyIndex%, Key$, sch$)
    lhKey& = MainKey + KeyIndex%
    lResult& = RegOpenKeyEx(lhKey&, Key, 0, KEY_SET_VALUE, lhKeyOpen&)
    If lResult& <> ERROR_SUCCESS Then Exit Sub
    lResult& = RegDeleteValue(lhKeyOpen&, sch$)
    'If lResult& <> ERROR_SUCCESS Then (Fehler...)
    RegCloseKey lhKeyOpen&
End Sub

Sub Reg_DeleteKey(KeyIndex%, Key$)
    lhKey& = MainKey + KeyIndex%
    lResult& = RegDeleteKey(lhKey&, Key$)
    'If lResult& <> ERROR_SUCCESS Then (Fehler...)
End Sub

Function Reg_Exist_Key(KeyIndex%, SubKey As String) As Boolean
    lhKey& = MainKey + KeyIndex
    Reg_Exist_Key = False
    l& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_ALL_ACCESS, lhKeyOpen&)
    'Schlüssel existiert nicht
    If l& <> ERROR_SUCCESS Then Exit Function
    Reg_Exist_Key = True
End Function

Function Reg_Exist_Value(KeyIndex%, SubKey As String, Name As String) As Boolean
    lhKey& = MainKey + KeyIndex
    Reg_Exist_Value = False
    l& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_ALL_ACCESS, lhKeyOpen&)
    'Schlüssel existiert nicht
    If l& <> ERROR_SUCCESS Then Exit Function
    'Wert existiert nicht
    l& = RegQueryValueExNULL(lhKeyOpen&, Name, 0&, lTyp&, 0&, cch&)
    If l& <> ERROR_SUCCESS Then Exit Function
    Reg_Exist_Value = True
End Function

Function Reg_GetValue_Typ(KeyIndex%, SubKey As String, Name As String) As String
    lhKey& = MainKey + KeyIndex
    Reg_GetValue_Typ = ""
    l& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_ALL_ACCESS, lhKeyOpen&)
    If l& <> ERROR_SUCCESS Then Exit Function
    
    l& = RegQueryValueExNULL(lhKeyOpen&, Name, 0&, lTyp&, 0&, cch&)
    If l& <> ERROR_SUCCESS Then Exit Function
    Select Case lTyp&
    Case REG_SZ
        Reg_GetValue_Typ = "STRING"
    Case REG_DWORD
        Reg_GetValue_Typ = "DWORD"
    Case REG_BINARY
        Reg_GetValue_Typ = "BINARY"
    Case Else
        Reg_GetValue_Typ = "???"
    End Select
End Function

Function Reg_GetValue(KeyIndex%, SubKey As String, Name As String) As String
    lhKey& = MainKey + KeyIndex
    Reg_GetValue = ""

    l& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_ALL_ACCESS, lhKeyOpen&): If l& <> ERROR_SUCCESS Then Exit Function
    l& = RegQueryValueExNULL(lhKeyOpen&, Name, 0&, lTyp&, 0&, cch&): If l& <> ERROR_SUCCESS Then Exit Function

    Select Case lTyp&
    Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ:
        sValue$ = String(cch& + 1, 0)
        l& = RegQueryValueExString(lhKeyOpen&, Name, 0&, lTyp&, sValue$, cch&): If l& <> ERROR_SUCCESS Then Exit Function
        Reg_GetValue = zeichennext$(Left$(sValue$, cch&), Chr$(0))
    Case REG_DWORD
        l& = RegQueryValueExLong(lhKeyOpen&, Name, 0&, lTyp&, lValue&, cch&): If l& <> ERROR_SUCCESS Then Exit Function
        Reg_GetValue = Trim$(Str$(lValue&))
    Case REG_BINARY
        sValue$ = String(cch& + 1, 0)
        l& = RegQueryValueExString(lhKeyOpen&, Name, 0&, lTyp&, sValue$, cch&): If l& <> ERROR_SUCCESS Then Exit Function
        Reg_GetValue = Left$(sValue$, cch&)
        For iTempInt = 1 To Len(Reg_GetValue)
            s = Asc(Mid$(Reg_GetValue, iTempInt, 1))
            'Binärwerte bringen Probleme,
            'sollte so aber funktionieren !
            Temp = ""
            If s = 26 Then Temp = "1A "
            If s = 58 Then Temp = "3A "
            If s = 74 Then Temp = "4A "
            If s = 90 Then Temp = "5A "
            If s = 106 Then Temp = "6A "
            If s = 122 Then Temp = "7A "
            If s = 138 Then Temp = "8A "
            If s = 154 Then Temp = "9A "
            If Temp = "" Then
                sBinaryString = sBinaryString & Format(Hex(Asc(Mid$(Reg_GetValue, iTempInt, 1))), "00") & " "
            Else
                sBinaryString = sBinaryString & Temp
            End If
        Next iTempInt
        Reg_GetValue = sBinaryString
    End Select
    RegCloseKey lhKeyOpen&
End Function


Function zeichennext$(a$, ch$)
    ai% = InStr(a$, ch$)
    If ai% = 0 Then
        zeichennext$ = a$: a$ = ""
    Else
        zeichennext$ = Left$(a$, ai% - 1): a$ = Mid$(a$, ai% + Len(ch$))
    End If
End Function

Sub RegCreateKey(KeyIndex As Integer, SubKey As String, NewSubKey As String)
    Dim Security As SECURITY_ATTRIBUTES
    lhKey& = MainKey + KeyIndex

    lResult& = RegOpenKeyEx(lhKey&, SubKey, 0, KEY_CREATE_SUB_KEY, lhKeyOpen&)
    If lResult& <> ERROR_SUCCESS Then Exit Sub

    lResult& = RegCreateKeyEx(lhKeyOpen&, NewSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew&, lDisposition&)
    If lResult& = ERROR_SUCCESS Then
        'If lDisposition& = REG_CREATED_NEW_KEY Then
            '   ...Schlüssel wurde angelegt
        'Else
            '   ...Schlüssel existiert bereits
        'End If
        RegCloseKey lhKeyNew&
    Else
        'Fehler...
    End If
    RegCloseKey lhKeyOpen&
End Sub

Sub Reg_SetBinary(MainKey%, Key$, sch$, wrt$)
    RegCreate MainKey%, Key$
    RegSetValue MainKey%, Key$, sch$, REG_BINARY, wrt$, Len(wrt$)
End Sub

Sub Reg_SetDWord(MainKey%, Key$, sch$, ByVal wrt&)
    RegCreate MainKey%, Key$
    w$ = ""
    For n% = 1 To Len(wrt&)
        w$ = w$ + Chr$(wrt& Mod 256)
        wrt& = Int(wrt& / 256)
    Next
    RegSetValue MainKey%, Key$, sch$, REG_DWORD, w$, Len(wrt&)
End Sub

Sub Reg_SetString(MainKey%, Key$, sch$, wrt$)
    RegCreate MainKey%, Key$
    RegSetValue MainKey%, Key$, sch$, REG_SZ, wrt$, Len(wrt$)
End Sub

