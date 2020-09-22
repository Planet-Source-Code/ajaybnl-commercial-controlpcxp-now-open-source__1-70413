Attribute VB_Name = "mod2"
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
Option Explicit
Public Enum ROOT_HKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_USERS = &H80000003
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_USERS, HKEY_LOCAL_MACHINE, HKEY_CURRENT_CONFIG
Private HKEY_DYN_DATA, HKEY_PERFORMANCE_DATA
#End If
Private HKEYS                                         As New Collection
Private REGTYPES                                      As New Collection
Private REGOUT                                        As New Collection
Public Const AddtoStartup_SZ                          As Long = 1
Private Const AddtoStartup_EXPAND_SZ                  As Long = 2
Private Const AddtoStartup_BINARY                     As Long = 3
Private Const AddtoStartup_DWORD                      As Long = 4
Private Const AddtoStartup_MULTI_SZ                   As Long = 7
Private Const AddtoStartup_OPTION_NON_VOLATILE        As Long = 0
Private Const AddtoStartup_CREATED_NEW_KEY            As Long = &H1
Private Const AddtoStartup_OPENED_EXISTING_KEY        As Long = &H2
Private Const KEY_QUERY_VALUE                         As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS                  As Long = &H8
Private Const KEY_NOTIFY                              As Long = &H10
Private Const READ_CONTROL                            As Long = &H20000
Private Const STANDARD_RIGHTS_ALL                     As Long = &H1F0000
Private Const STANDARD_RIGHTS_READ                    As Long = (READ_CONTROL)
Private Const SYNCHRONIZE                             As Long = &H100000
Private Const KEY_READ                                As Double = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_SET_VALUE                           As Long = &H2
Private Const KEY_CREATE_SUB_KEY                      As Long = &H4
Private Const KEY_CREATE_LINK                         As Long = &H20
Private Const STANDARD_RIGHTS_WRITE                   As Long = (READ_CONTROL)
Private Const KEY_WRITE                               As Double = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS                          As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS                            As Long = 0
Type SECURITY_ATTRIBUTES
    nLength                                             As Long
    lpSecurityDescriptor                                As Long
    bInheritHandle                                      As Boolean
End Type
Public Type FILETIME
    dwLowDateTime                                       As Long
    dwHighDateTime                                      As Long
End Type
Public Type KEYARRAY
    cnt                                                 As Long
    key()                                               As String
    Data()                                              As Variant
    DataType()                                          As Long
    DataSize()                                          As Long
End Type
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, dwSize As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal dwReserved As Long, ByVal dwType As Long, lpValue As Any, ByVal dwSize As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Function DeleteRegValue(hKey As Long, subkey As String, ValueName As String) As Long
Dim Result     As Long
Dim hKeyResult As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    Result = RegOpenKeyEx(hKey, subkey, 0, KEY_WRITE, hKeyResult)
    If Result <> ERROR_SUCCESS Then
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 1)
'<:-) Convert 'If..Then/Exit/End If'  to
'<:-) 'If Not .. Then/Rest_Of_Code/End If'
    End If
    Result = RegDeleteValue(hKeyResult, ValueName)
    DeleteRegValue = Result
    RegCloseKey hKeyResult
End Function
Public Function GetRegKey(ByVal hKeyRoot As String)
    Select Case hKeyRoot
    Case "HKEY_CLASSES_ROOT"
        GetRegKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetRegKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetRegKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetRegKey = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetRegKey = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetRegKey = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetRegKey = HKEY_DYN_DATA
    End Select
End Function
Public Function GetRegKeyName(ByVal hKeyRoot As Long) As String
    Select Case hKeyRoot
    Case HKEY_CLASSES_ROOT
        GetRegKeyName = "HKEY_CLASSES_ROOT"
    Case HKEY_CURRENT_USER
        GetRegKeyName = "HKEY_CURRENT_USER"
    Case HKEY_LOCAL_MACHINE
        GetRegKeyName = "HKEY_LOCAL_MACHINE"
    Case HKEY_USERS
        GetRegKeyName = "HKEY_USERS"
    Case HKEY_PERFORMANCE_DATA
        GetRegKeyName = "HKEY_PERFORMANCE_DATA"
    Case HKEY_CURRENT_CONFIG
        GetRegKeyName = "HKEY_CURRENT_CONFIG"
    Case HKEY_DYN_DATA
        GetRegKeyName = "HKEY_DYN_DATA"
    End Select
End Function
Public Function GetSubKeys(hKey As Long, subkey As String) As KEYARRAY
Dim lResult   As Long
Dim ft        As FILETIME
Dim SubKeyCnt As Long
Dim hSubKey   As Long
Dim i         As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    lResult = RegOpenKeyEx(hKey, subkey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)
    If lResult <> ERROR_SUCCESS Then
        GetSubKeys.cnt = 0
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    lResult = RegQueryInfoKey(hSubKey, vbNullString, 0, 0, SubKeyCnt, 65, 0, 0, 0, 0, 0, ft)
    If (lResult <> ERROR_SUCCESS) Or (SubKeyCnt <= 0) Then
        GetSubKeys.cnt = 0
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    With GetSubKeys
        .cnt = SubKeyCnt
        ReDim .key(SubKeyCnt + 1)
        ReDim .Data(SubKeyCnt + 1)
'ReDim GetSubKeys. (SubKeyCnt + 1)
    End With 'GetSubKeys
    For i = 0 To SubKeyCnt - 1
        With GetSubKeys
            .key(i) = String$(65, 0)
            RegEnumKeyEx hSubKey, i, .key(i), 65, 0, vbNullString, 0, ft
            .key(i) = StripNulls(.key(i))
            Call RegKeysCallback(.key(i))
        End With
    Next i
    RegCloseKey hSubKey
End Function
Private Sub InitReg()
    REGTYPES.Add 1, "AddtoStartup_SZ"
    REGOUT.Add 1, ""
    REGTYPES.Add 2, "AddtoStartup_EXPAND_SZ"
    REGOUT.Add 2, "hex(2):"
    REGTYPES.Add 3, "AddtoStartup_BINARY"
    REGOUT.Add 3, "hex:"
    REGTYPES.Add 4, "AddtoStartup_DWORD"
    REGOUT.Add 4, "dword:"
    With HKEYS
        .Add ROOT_HKEY.HKEY_CLASSES_ROOT, "HKEY_CLASSES_ROOT"
        .Add ROOT_HKEY.HKEY_CURRENT_USER, "HKEY_CURRENT_USER"
        .Add ROOT_HKEY.HKEY_USERS, "HKEY_USERS"
        .Add ROOT_HKEY.HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE"
        .Add ROOT_HKEY.HKEY_CURRENT_CONFIG, "HKEY_CURRENT_CONFIG"
        .Add ROOT_HKEY.HKEY_DYN_DATA, "HKEY_DYN_DATA"
        .Add ROOT_HKEY.HKEY_PERFORMANCE_DATA, "HKEY_PERFORMANCE_DATA"
    End With
End Sub
Public Function ReadReg(hKey As Long, subkey As String, DataName As String, DefaultData As Variant) As Variant
Dim hKeyResult As Long
Dim lData      As Long
Dim sData      As String
Dim DataType   As Long
Dim DataSize   As Long
Dim Result     As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    ReadReg = DefaultData
    Result = RegOpenKeyEx(hKey, subkey, 0, KEY_QUERY_VALUE, hKeyResult)
    If Result <> ERROR_SUCCESS Then
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 1)
'<:-) Convert 'If..Then/Exit/End If'  to
'<:-) 'If Not .. Then/Rest_Of_Code/End If'
    End If
    Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, ByVal 0, DataSize)
    If Result <> ERROR_SUCCESS Then
        RegCloseKey hKeyResult
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    Select Case DataType
    Case AddtoStartup_SZ
        sData = Space$(DataSize + 1)
        Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, ByVal sData, DataSize)
        If Result = ERROR_SUCCESS Then
            ReadReg = CVar(StripNulls(RTrim$(sData)))
        End If
    Case AddtoStartup_DWORD
        Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, lData, 4)
        If Result = ERROR_SUCCESS Then
            ReadReg = CVar(lData)
        End If
    End Select
    RegCloseKey hKeyResult
End Function
Public Function RegGetValues(hKey As Long, subkey As String) As KEYARRAY
Dim hSubKey         As Long
Dim i               As Long
Dim s               As String
Dim lResult         As Long
Dim ValName         As String
Dim ValSize         As Long
Dim LastWriteTime   As FILETIME
Dim SubKeyCnt       As Long
Dim MaxSubKeyLen    As Long
Dim ValueCnt        As Long
Dim MaxValueNameLen As Long
Dim MaxValueLen     As Long
Dim SecurityDesc    As Long
Dim DataType        As Long
Dim DataSize        As Long
Dim ba()            As Byte
    If HKEYS.Count = 0 Then
        InitReg
    End If
    lResult = RegOpenKeyEx(hKey, subkey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)
    If lResult <> ERROR_SUCCESS Then
        RegGetValues.cnt = 0
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    lResult = RegQueryInfoKey(hSubKey, vbNull, 0, 0, SubKeyCnt, MaxSubKeyLen, vbNull, ValueCnt, MaxValueNameLen, MaxValueLen, SecurityDesc, LastWriteTime)
    RegGetValues.cnt = 0
    ValSize = MaxValueNameLen + 100
    ValName = String$(ValSize + 1, 0)
    ReDim ba(MaxValueLen + 1) As Byte
    DataType = 0
    DataSize = 0
    ValSize = MaxValueNameLen + 100
    ValName = String$(ValSize + 1, 0)
    DataSize = UBound(ba) - 1
    lResult = RegEnumValue(hSubKey, RegGetValues.cnt, ValName, Len(ValName), 0, DataType, ba(0), DataSize)
    Do While lResult = ERROR_SUCCESS
        With RegGetValues
            .cnt = .cnt + 1
            ReDim Preserve .key(.cnt + 1)
            ReDim Preserve .DataType(.cnt + 1)
            ReDim Preserve .DataSize(.cnt + 1)
            ReDim Preserve .Data(.cnt + 1)
            .key(.cnt) = StripNulls(ValName)
            .DataType(.cnt) = DataType
        End With
        Select Case DataType
        Case AddtoStartup_SZ
            s = ""
            i = 0
            Do While i < DataSize + 1
                s = s & Chr$(ba(i))
                i = i + 1
            Loop
            RegGetValues.Data(RegGetValues.cnt) = StripNulls(s)
            RegGetValues.DataSize(RegGetValues.cnt) = Len(s)
        Case AddtoStartup_EXPAND_SZ
            s = ""
            i = 0
            Do While i < (DataSize * 2) + 1
                s = s & Chr$(ba(i))
                i = i + 1
            Loop
            RegGetValues.Data(RegGetValues.cnt) = s
            RegGetValues.DataSize(RegGetValues.cnt) = (DataSize * 2)
        Case AddtoStartup_MULTI_SZ
            s = ""
            i = 0
            Do While i < (DataSize * 2) + 1
                s = s & Chr$(ba(i))
                i = i + 1
            Loop
            RegGetValues.Data(RegGetValues.cnt) = s
            RegGetValues.DataSize(RegGetValues.cnt) = (DataSize * 2)
        Case AddtoStartup_BINARY
            RegGetValues.Data(RegGetValues.cnt) = ba
            RegGetValues.DataSize(RegGetValues.cnt) = DataSize
        Case AddtoStartup_DWORD
            i = ba(0)
            i = i + (CLng(ba(1)) * 256)
            i = i + (CLng(ba(2)) * 256 * 256)
            s = Hex$(i)
            If Len(s) < 6 Then
                s = String$(6 - Len(s), "0") & s
            End If
            s = "&h" & Hex$(ba(3)) & s
            RegGetValues.Data(RegGetValues.cnt) = Val(s)
            RegGetValues.DataSize(RegGetValues.cnt) = 4
        Case Else
            RegGetValues.Data(RegGetValues.cnt) = ba
            RegGetValues.DataSize(RegGetValues.cnt) = DataSize
        End Select
        DataType = 0
        DataSize = UBound(ba) - 1
        ValSize = MaxValueNameLen + 100
        ValName = String$(ValSize + 1, 0)
        lResult = RegEnumValue(hSubKey, RegGetValues.cnt, ValName, Len(ValName), 0, DataType, ba(0), DataSize)
    Loop
    RegCloseKey hSubKey
End Function
Private Function StripNulls(ByVal s As String) As String
Dim i As Long
    i = InStr(s, vbNullChar)
    If i > 0 Then
        StripNulls = Left$(s, i - 1)
    Else
        StripNulls = s
    End If
End Function
Public Function WriteRegString(hKey As Long, subkey As String, DataName As String, DataValue As String) As Long
Dim sa           As SECURITY_ATTRIBUTES
Dim hKeyResult   As Long
Dim lDisposition As Long
Dim Result       As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    With sa
        .nLength = Len(sa)
        .lpSecurityDescriptor = 0
        .bInheritHandle = False
    End With
    Result = RegCreateKeyEx(hKey, subkey, 0, vbNullString, AddtoStartup_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sa, hKeyResult, lDisposition)
    If DataValue <= "" Then
        DataValue = ""
    End If
    If (Result = ERROR_SUCCESS) Or (Result = AddtoStartup_CREATED_NEW_KEY) Or (Result = AddtoStartup_OPENED_EXISTING_KEY) Then
        Result = RegSetValueEx(hKeyResult, DataName, 0&, AddtoStartup_SZ, ByVal DataValue, Len(DataValue))
        RegCloseKey hKeyResult
    End If
    WriteRegString = Result
End Function
''
''Public Sub ServiceCreator(ByVal strName As String, ByVal strPath As String)
''
''
''
''
''
''
''Dim hKey         As Long                      ' receives handle to the registry key
''Dim secattr      As SECURITY_ATTRIBUTES       ' security settings for the key
''Dim subkey       As String                ' name of the subkey to create or open
''Dim neworused    As Long                 ' receives flag for if the key was created or opened
''Dim stringbuffer As String          ' the string to put into the registry
''Dim retval       As Long                    ' return value
''' Set the name of the new key and the default security settings
''subkey = "System\CurrentControlSet\Services\" & strName & "\Parameters"
''
''With secattr
''.nLength = Len(secattr)
''.lpSecurityDescriptor = 0
''.bInheritHandle = 1
''' Create (or open) the registry key.
''End With 'secattr
''retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, vbNullString, 0, KEY_WRITE, secattr, hKey, neworused)
''
''
''If retval <> 0 Then
'''Debug.Print "Error opening or creating registry key -- aborting."
''Exit Sub
''End If
''' Write the string to the registry.  Note the use of ByVal in the second-to-last
''' parameter because we are passing a string.
''stringbuffer = strPath & vbNullChar    ' the terminating null is necessary
''retval = RegSetValueEx(hKey, "Application", 0, AddtoStartup_SZ, ByVal stringbuffer, Len(stringbuffer))
''' Close the registry key.
''retval = RegCloseKey(hKey)
''End Sub
''
':)Code Fixer V3.0.9 (11/15/2006 12:12:17 PM) 67 + 378 = 445 Lines Thanks Ulli for inspiration and lots of code.
